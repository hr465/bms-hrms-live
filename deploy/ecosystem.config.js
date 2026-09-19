// pm2 process definition. Settings come from /home/ubuntu/hrms/.env (never committed).
module.exports = {
  apps: [{
    name: "hrms",
    script: "src/server.js",
    cwd: __dirname + "/..",
    node_args: "--env-file=.env",
    max_memory_restart: "700M",
    autorestart: true,
  }],
};
