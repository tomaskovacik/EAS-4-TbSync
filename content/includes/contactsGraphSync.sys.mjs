// Contacts delta sync using Microsoft Graph
export const contactsGraphSync = {
  async syncContacts(syncData) {
    const accountData = syncData.accountData;
    if (!("graphClient" in eas) || !("graphMappings" in eas)) throw new Error("Graph modules missing");
    const graphClient = eas.graphClient;
    const graphMappings = eas.graphMappings;
    const deltaKey = "graph.contacts.deltaLink";
    let deltaLink = accountData.getAccountProperty(deltaKey) || "";

    try {
      let response = await graphClient.contactsDelta(accountData, deltaLink || null);
      while (response) {
        const entries = response.value || [];
        for (const entry of entries) {
          if (entry["@removed"]) {
            const remoteId = entry.id || (entry["@removed"] && entry["@removed"].id) || null;
            await eas.sync.applyRemoteContact(syncData, null, "delete", { remoteId });
          } else {
            // Map Graph contact -> vCard-like -> EAS data
            const vcard = graphMappings.contactGraphToVCard(entry);
            await eas.sync.applyRemoteContact(syncData, vcard, "upsert", { remoteId: vcard.remoteId || entry.id });
          }
        }
        if (response["@odata.deltaLink"]) {
          accountData.setAccountProperty(deltaKey, response["@odata.deltaLink"]);
          break;
        }
        if (response["@odata.nextLink"]) {
          response = await graphClient.contactsDelta(accountData, response["@odata.nextLink"]);
        } else break;
      }
    } catch (e) {
      console.error("contactsGraphSync error", e);
      throw e;
    }
  }
};