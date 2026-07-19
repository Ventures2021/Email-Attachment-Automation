const { AttachmentAutomation, AuditLog, JsonStateStore, MicrosoftGraphClient, NotionClient, loadConfig } = require('./src/automation');

async function main() {
  const config = loadConfig();
  const graphClient = new MicrosoftGraphClient(config);
  const notionClient = new NotionClient(config);
  const stateStore = new JsonStateStore(config.stateFile);
  const auditLog = new AuditLog(config.auditLogFile);
  const automation = new AttachmentAutomation({
    auditLog,
    config,
    graphClient,
    notionClient,
    stateStore,
  });

  if (config.runOnce) {
    await automation.runOnce();
    return;
  }

  await automation.start();
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
