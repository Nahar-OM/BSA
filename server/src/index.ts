import { Hono } from 'hono';
import { cors } from 'hono/cors';
import { streamSSE } from 'hono/streaming';
import { spawn } from 'child_process';
import { serveStatic } from 'hono/bun';
import { join, basename } from 'path';
import { readFile, readdir, unlink, rmdir, stat } from 'fs/promises';
import { getSignedUrl } from '@aws-sdk/s3-request-presigner';
import { PutObjectCommand } from '@aws-sdk/client-s3';
import { s3Client } from './config/s3-init';

const app = new Hono();

app.use('/*', cors());

app.use('/out/*', serveStatic({ root: './' }));

async function deleteOutFolderContents(dir: string) {
  const files = await readdir(dir);

  for (const file of files) {
    const filePath = join(dir, file);
    const fileStat = await stat(filePath);

    if (fileStat.isDirectory()) {
      await deleteOutFolderContents(filePath);
      await rmdir(filePath);
    } else {
      await unlink(filePath);
    }
  }

  console.log('Contents of "out" folder deleted.');
}

function runPythonScript(scriptPath: string, dataFolderName: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const process = spawn('/usr/local/bin/python3', [scriptPath, dataFolderName]);
    let output = '';
    let errorOutput = '';

    process.stdout.on('data', (data) => {
      output += data.toString();
      console.log(`Python script output: ${data}`);
    });

    process.stderr.on('data', (data) => {
      errorOutput += data.toString();
      console.error(`Python script error: ${data}`);
    });

    process.on('close', (code) => {
      if (code === 0) {
        resolve(output.trim());
      } else {
        reject(`Python script exited with code ${code}. Error: ${errorOutput}`);
      }
    });
  });
}

async function uploadFileToS3(filePath: string): Promise<string> {
  const fileContent = await readFile(filePath);
  const fileName = `${crypto.randomUUID()}_Report.xlsx`;

  const command = new PutObjectCommand({
    Bucket: process.env.AWS_S3_BUCKET_NAME,
    Key: fileName,
    Body: fileContent,
    ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  const signedUrl = await getSignedUrl(s3Client, command, { expiresIn: 3600 });

  await fetch(signedUrl, {
    method: 'PUT',
    body: fileContent,
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
  });

  const reportUrl = `https://${process.env.AWS_S3_BUCKET_NAME}.s3.amazonaws.com/${fileName}`;
  return reportUrl;
}


app.get('/run-bsa', async (c) => {
  const dataFolderName = c.req.query('folder');

  if (!dataFolderName) {
    return c.text("Missing 'folder' parameter in the URL", 400);
  }

  return streamSSE(c, async (stream) => {
    await stream.writeSSE({ data: "BSA process started. This may take a while..." });

    try {
      const result = await runPythonScript('../ML/index.py', dataFolderName);
      await stream.writeSSE({ data: `BSA process completed. Result: ${result || "BSA process completed successfully"}` });

      const filePath = join(process.cwd(), 'out', 'report_files', 'Main_Report.xlsx');

      console.log(`Attempting to access file at: ${filePath}`);

      if (!await Bun.file(filePath).exists()) {
        throw new Error(`File not found: ${filePath}`);
      }

      const signedUrl = await uploadFileToS3(filePath);
      await stream.writeSSE({ data: `Download URL: ${signedUrl}` });
      await deleteOutFolderContents(join(process.cwd(), 'out'));
    } catch (error) {
      console.error('Error in BSA process:', error);
      await stream.writeSSE({ data: `Error: ${error}` });
    } finally {
      await stream.close();
    }
  });
});

// Route for downloading files
app.get('/download/:filename*', async (c) => {
  const filename = c.req.param('filename');
  const filePath = join(process.cwd(), 'out', filename!);

  const file = Bun.file(filePath);
  const exists = await file.exists();

  if (!exists) {
    return c.notFound();
  }

  return new Response(file, {
    headers: {
      "Content-Type": file.type,
      "Content-Disposition": `attachment; filename="${basename(filePath)}"`,
    },
  });
});

// Simple root route
app.get('/', (c) => c.text('Hello World!'));

export default app;
