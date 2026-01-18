// Exports contactsGraphSync
// Uses eas.graphClient and eas.graphMappings to fetch contact deltas and map them.
// After mapping it calls eas.sync.applyRemoteContact(syncData, vcard, changeType)
// which is a small adaptor hook. Implement local storage writes there (TODO).

export const contactsGraphSync = {
  // Main entry point called by sync.js
  async syncContacts(syncData) {
    const accountData = syncData.accountData;

    // Ensure graph client exists
    if (!("graphClient" in eas) || !("graphMappings" in eas)) {
      throw new Error("Graph modules are not loaded (eas.graphClient / eas.graphMappings missing).");
    }

    const graphClient = eas.graphClient;
    const graphMappings = eas.graphMappings;

    // Load previously saved deltaLink (if any)
    const deltaKey = "graph.contacts.deltaLink";
    let deltaLink = accountData.getAccountProperty(deltaKey) || "";

    // Iterate through pages until we receive @odata.deltaLink
    let next = null;
    try {
      let response = await graphClient.contactsDelta(accountData, deltaLink || null);

      // response should contain .value array and maybe @odata.nextLink / @odata.deltaLink
      while (response) {
        const entries = response.value || [];
        for (const entry of entries) {
          if (entry["@removed"]) {
            // deletion - Graph signals removed items with @removed
            const remoteId = entry.id || (entry["@removed"] && entry["@removed"].id) || null;
            // Call adapter to apply deletion locally
            await this._applyChange(syncData, null, remoteId, "delete");
          } else {
            // create or update
            const vcard = graphMappings.contactGraphToVCard(entry);
            await this._applyChange(syncData, vcard, vcard.remoteId || entry.id, "upsert");
          }
        }

        // deltaLink received -> we're done; persist it
        if (response["@odata.deltaLink"]) {
          const newDelta = response["@odata.deltaLink"];
          accountData.setAccountProperty(deltaKey, newDelta);
          break;
        }

        // nextLink -> fetch next page
        if (response["@odata.nextLink"]) {
          response = await graphClient.contactsDelta(accountData, response["@odata.nextLink"]);
        } else {
          // no next / delta link -> stop
          break;
        }
      } // end while
    } catch (e) {
      // Bubble up error so the sync framework can handle retry/backoff.
      // Keep messages readable for log/telemetry
      console.error("contactsGraphSync.syncContacts error:", e);
      throw e;
    }
  },

  // Internal helper: route a single change
  async _applyChange(syncData, vcard, remoteId, changeType) {
    // changeType is "delete" or "upsert"
    // This function defers actual local apply logic to an adapter hook so we don't
    // duplicate address-book logic here. Implement the adapter in one of two ways:
    //  - Add an implementation to eas.sync.applyRemoteContact(syncData, vcard, changeType)
    //  - Or extend this function to call addressbook APIs directly.
    //
    // We prefer the adapter approach to keep responsibilities separated.
    if (typeof eas.sync.applyRemoteContact === "function") {
      return await eas.sync.applyRemoteContact(syncData, vcard, changeType, { remoteId });
    } else {
      // Fallback: basic logging only (no-local write).
      console.log("Graph contact change:", changeType, remoteId, vcard ? vcard.fn : "(deleted)");
      return;
    }
  }
};