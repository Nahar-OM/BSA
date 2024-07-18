import { Hono } from 'hono'
import { cors } from 'hono/cors'
import { streamSSE } from 'hono/streaming'
import { spawn } from 'child_process'
import { serveStatic } from 'hono/bun'
import { join, basename } from 'path'

const app = new Hono()

// Add CORS middleware
app.use('/*', cors())

app.use('/public/*', serveStatic({ root: './' }))

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

      // Assuming the Python script returns the relative path of the output
      const outputPath = result.trim();
      const downloadUrl = `/download/${basename(outputPath)}`;
      await stream.writeSSE({ data: `Download URL: ${downloadUrl}` });
    } catch (error) {
      console.error('Error in BSA process:', error);
      await stream.writeSSE({ data: `Error: ${error}` });
    } finally {
      // Ensure the connection is closed
      await stream.close();
    }
  });
})

// Route for downloading files
app.get('/download/:filename', async (c) => {
  const filename = c.req.param('filename');
  const filePath = join(process.cwd(), 'public', 'out', filename);

  const file = Bun.file(filePath);
  const exists = await file.exists();

  if (!exists) {
    return c.notFound();
  }

  return new Response(file, {
    headers: {
      "Content-Type": file.type,
      "Content-Disposition": `attachment; filename="${filename}"`,
    },
  });
})

app.get('/', (c) => c.text('Hello World!'))

export default app
