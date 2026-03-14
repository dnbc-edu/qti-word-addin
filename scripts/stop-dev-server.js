import { exec } from 'node:child_process';

function execCommand(command) {
  return new Promise((resolve, reject) => {
    exec(command, (error, stdout, stderr) => {
      if (error) {
        reject({
          code: error.code,
          message: stderr || error.message
        });
        return;
      }

      resolve(stdout);
    });
  });
}

async function getPidsOnPort(port) {
  let output = '';

  try {
    output = await execCommand(`lsof -ti tcp:${port}`);
  } catch (error) {
    if (error.code === 1) {
      return [];
    }

    throw error;
  }

  return output
    .split('\n')
    .map((value) => value.trim())
    .filter(Boolean);
}

async function killPid(pid) {
  await execCommand(`kill ${pid}`);
}

async function main() {
  const port = Number(process.env.DEV_SERVER_PORT || 3000);

  try {
    const pids = await getPidsOnPort(port);

    if (!pids.length) {
      console.log(`No process found on port ${port}.`);
      return;
    }

    for (const pid of pids) {
      await killPid(pid);
    }

    console.log(`Stopped ${pids.length} process(es) on port ${port}: ${pids.join(', ')}`);
  } catch (error) {
    console.error(`Failed to stop dev server: ${error.message || String(error)}`);
    process.exit(1);
  }
}

main();
