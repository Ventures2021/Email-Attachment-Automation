const { runSimpleIntakeCli } = require('./src/simple-intake');

runSimpleIntakeCli(process.argv).catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
