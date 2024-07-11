import { spawn } from 'child_process';
import { serve } from "bun";

function runPythonScript(scriptPath: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const process = spawn('/usr/local/bin/python3', [scriptPath]);
    let output = '';

    process.stdout.on('data', (data) => {
      output += data.toString();
      console.log(`Python script output: ${data}`);
    });

    process.stderr.on('data', (data) => {
      console.error(`Python script error: ${data}`);
    });

    process.on('close', (code) => {
      if (code === 0) {
        resolve(output.trim());
      } else {
        reject(`Python script exited with code ${code}`);
      }
    });
  });
}

const server = serve({
  port: 3000,
  async fetch(req) {
    const url = new URL(req.url);

    if (url.pathname === "/run-bsa") {
      const stream = new ReadableStream({
        async start(controller) {
          controller.enqueue("BSA process started. This may take a while...\n");

          try {
            const result = await runPythonScript('./index.py');
            controller.enqueue(`BSA process completed. Result: ${result || "BSA process completed successfully"}\n`);
          } catch (error) {
            controller.enqueue(`Error: ${error}\n`);
          }

          controller.close();
        }
      });

      return new Response(stream, {
        headers: { "Content-Type": "text/plain" },
      });
    }

    return new Response("Hello World!");
  },
});

console.log(`Listening on http://localhost:${server.port}`);
