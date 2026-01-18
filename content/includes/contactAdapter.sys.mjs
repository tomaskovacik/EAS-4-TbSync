// Implements eas.sync.applyRemoteContact and eas.sync.applyRemoteEvent using existing TB helpers.
// This module reuses contactsync.setThunderbirdItemFromWbxml and calendar write helpers.

ChromeUtils.defineESModuleGetters(this, {
  AddrBookCard: "resource:///modules/AddrBookCard.sys.mjs",
  VCardUtils: "resource:///modules/VCardUtils.sys.mjs",
});

export const contactAdapter = {
  // Apply contact change from Graph: vcard is vCard-like object produced by graphMappings.contactGraphToVCard
  async applyRemoteContact(syncData, vcard, changeType, meta = {}) {
    const accountData = syncData.accountData;
    const remoteId = meta.remoteId || (vcard && vcard.remoteId) || "";
    // Build minimal EAS-style ApplicationData object using mapping helper
    const easData = (vcard && eas.graphMappings.contactToEasData(vcard)) || {};

    // Use existing syncData.target API to find/add/modify/delete items
    try {
      if (changeType === "delete") {
        const found = await syncData.target.getItem(remoteId);
        if (found) {
          // The provider target API expects deletion by server id - use deleteItem if available
          if (typeof syncData.target.deleteItem === "function") {
            await syncData.target.deleteItem(remoteId);
          } else if (typeof syncData.target.removeItem === "function") {
            await syncData.target.removeItem(found);
          } else {
            // fallback: mark for deletion via modify with tombstone if supported
            console.log("delete API not found on target; skipping delete for", remoteId);
          }
        }
        return;
      }

      // For upsert: if item exists -> modify; otherwise add
      let foundItem = await syncData.target.getItem(remoteId);
      if (!foundItem) {
        // create new Thunderbird item via existing create flow
        const newItem = eas.sync.createItem(syncData);
        // contactsync.setThunderbirdItemFromWbxml expects an EAS ApplicationData object
        if (eas && eas.sync && eas.sync.contacts && typeof eas.sync.contacts.setThunderbirdItemFromWbxml === "function") {
          eas.sync.contacts.setThunderbirdItemFromWbxml(newItem, easData, remoteId, syncData, "graph");
        } else {
          // fallback: try applying minimal fields directly on card
          if (newItem._card && newItem._card.vCardProperties) {
            newItem._card.displayName = easData.FileAs || "";
            // primary email
            if (easData.Email1Address) newItem._card.primaryEmail = easData.Email1Address;
          }
        }
        await syncData.target.addItem(newItem);
      } else {
        // update existing item
        const updated = foundItem.clone();
        if (eas && eas.sync && eas.sync.contacts && typeof eas.sync.contacts.setThunderbirdItemFromWbxml === "function") {
          eas.sync.contacts.setThunderbirdItemFromWbxml(updated, easData, remoteId, syncData, "graph");
          await syncData.target.modifyItem(updated, foundItem);
        } else {
          // minimal update: set displayName / primaryEmail
          if (updated._card) {
            updated._card.displayName = easData.FileAs || updated._card.displayName;
            if (easData.Email1Address) updated._card.primaryEmail = easData.Email1Address;
          }
          await syncData.target.modifyItem(updated, foundItem);
        }
      }
    } catch (e) {
      console.error("applyRemoteContact error:", e);
      throw e;
    }
  },

  // Apply calendar event change from Graph (data is EAS-like ApplicationData produced by graphMappings.eventGraphToEasData)
  async applyRemoteEvent(syncData, data, changeType, meta = {}) {
    const remoteId = meta.remoteId || (data && data._graphId) || "";
    try {
      if (changeType === "delete") {
        // remove by server id if API exists
        if (typeof syncData.target.deleteItem === "function") {
          await syncData.target.deleteItem(remoteId);
        } else {
          const found = await syncData.target.getItem(remoteId);
          if (found && typeof syncData.target.removeItem === "function") await syncData.target.removeItem(found);
          else console.log("No delete API for calendar target; skipping", remoteId);
        }
        return;
      }

      // upsert logic
      const foundItem = await syncData.target.getItem(remoteId);
      if (!foundItem) {
        const newItem = eas.sync.createItem(syncData);
        // Attempt to leverage calendar write helpers if present
        if (eas && eas.sync && eas.sync.Calendar && typeof eas.sync.Calendar.setThunderbirdItemFromWbxml === "function") {
          eas.sync.Calendar.setThunderbirdItemFromWbxml(newItem, data, remoteId, syncData, "graph");
        } else if (typeof eas.sync.setItemSubject === "function") {
          // Lightweight mapping
          eas.sync.setItemSubject(newItem, syncData, { Subject: data.Subject });
          eas.sync.setItemBody(newItem, syncData, { Body: data.Body || "" });
        }
        await syncData.target.addItem(newItem);
      } else {
        const updated = foundItem.clone();
        if (eas && eas.sync && eas.sync.Calendar && typeof eas.sync.Calendar.setThunderbirdItemFromWbxml === "function") {
          eas.sync.Calendar.setThunderbirdItemFromWbxml(updated, data, remoteId, syncData, "graph");
          await syncData.target.modifyItem(updated, foundItem);
        } else {
          if (typeof eas.sync.setItemSubject === "function") {
            eas.sync.setItemSubject(updated, syncData, { Subject: data.Subject });
            eas.sync.setItemBody(updated, syncData, { Body: data.Body || "" });
          }
          await syncData.target.modifyItem(updated, foundItem);
        }
      }
    } catch (e) {
      console.error("applyRemoteEvent error:", e);
      throw e;
    }
  }
};