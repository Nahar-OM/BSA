import { spawn } from 'child_process';
import { serve } from "bun";

function runPythonScript(scriptPath: string, args: string[]) {
  return new Promise((resolve, reject) => {
    const process = spawn('/usr/local/bin/python', [scriptPath, ...args]);

    process.stdout.on('data', (data) => {
      console.log(`Python script output: ${data}`);
    });

    process.stderr.on('data', (data) => {
      console.error(`Python script error: ${data}`);
    });

    process.on('close', (code) => {
      if (code === 0) {
        resolve('Python script executed successfully');
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
      try {
        await runPythonScript('./BSA_main.py', ['Main_BSA_Function', './bank-statement/LANDCRAFT-RECREATIONS', 'LANDCRAFT RECREATIONS']);
        return new Response("BSA process completed successfully");
      } catch (error) {
        return new Response(`Error: ${error}`, { status: 500 });
      }
    }

    return new Response("Hello World!");
  },
});

console.log(`Listening on http://localhost:${server.port}`);
