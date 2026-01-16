const { spawn } = require("child_process");

const processes = [
  {
    name: "services-proxy",
    command: "node",
    args: ["services-proxy-server/src/index.js"],
    env: {
      ...process.env,
      SERVICES_PROXY_PORT: process.env.SERVICES_PROXY_PORT ?? "4100"
    }
  },
  {
    name: "agent-proxy",
    command: "node",
    args: ["agent-proxy-server/src/index.js"],
    env: { ...process.env, PORT: process.env.AGENT_PROXY_PORT ?? process.env.PORT ?? "4000" }
  }
];

let shuttingDown = false;

const children = processes.map(({ name, command, args, env }) => {
  const child = spawn(command, args, { stdio: "inherit", env });
  child.on("exit", (code, signal) => {
    if (signal) {
      console.log(`${name} exited due to signal ${signal}`);
    } else {
      console.log(`${name} exited with code ${code}`);
    }
    if (!shuttingDown) {
      shuttingDown = true;
      terminateChildren();
      process.exit(code ?? 1);
    }
  });
  return child;
});

function terminateChildren() {
  children.forEach((proc) => {
    if (!proc.killed) {
      proc.kill("SIGTERM");
    }
  });
}

function shutdown(signal) {
  if (shuttingDown) {
    return;
  }
  shuttingDown = true;
  console.log(`Received ${signal}. Shutting down child processes...`);
  terminateChildren();
}

process.on("SIGINT", () => shutdown("SIGINT"));
process.on("SIGTERM", () => shutdown("SIGTERM"));
