// Calendars delta sync using Microsoft Graph
export const calendarsGraphSync = {
  // Sync all calendars (or a specific calendar) via delta queries.
  // If calendarId param is provided, sync only that calendar; otherwise iterate calendars.
  async syncCalendars(syncData, calendarId = null) {
    const accountData = syncData.accountData;
    if (!("graphClient" in eas) || !("graphMappings" in eas)) throw new Error("Graph modules missing");
    const graphClient = eas.graphClient;
    const graphMappings = eas.graphMappings;

    // If calendarId not provided, enumerate calendars and sync each
    if (!calendarId) {
      const calendarsResp = await graphClient.getCalendars(accountData);
      const calendars = (calendarsResp && calendarsResp.value) || [];
      for (const cal of calendars) {
        await this._syncSingleCalendar(syncData, cal.id);
      }
      return;
    } else {
      await this._syncSingleCalendar(syncData, calendarId);
    }
  },

  async _syncSingleCalendar(syncData, calendarId) {
    const accountData = syncData.accountData;
    const graphClient = eas.graphClient;
    const graphMappings = eas.graphMappings;
    const deltaKey = `graph.calendar.deltaLink.${calendarId}`;
    let deltaLink = accountData.getAccountProperty(deltaKey) || "";

    try {
      let response = await graphClient.eventsDelta(accountData, calendarId, deltaLink || null);
      while (response) {
        const entries = response.value || [];
        for (const entry of entries) {
          // Graph returns deleted items with @removed
          if (entry["@removed"]) {
            const remoteId = entry.id || (entry["@removed"] && entry["@removed"].id) || null;
            await eas.sync.applyRemoteEvent(syncData, null, "delete", { remoteId, calendarId });
          } else {
            const data = graphMappings.eventGraphToEasData(entry);
            await eas.sync.applyRemoteEvent(syncData, data, "upsert", { remoteId: entry.id, calendarId });
          }
        }
        if (response["@odata.deltaLink"]) {
          accountData.setAccountProperty(deltaKey, response["@odata.deltaLink"]);
          break;
        }
        if (response["@odata.nextLink"]) {
          response = await graphClient.eventsDelta(accountData, calendarId, response["@odata.nextLink"]);
        } else break;
      }
    } catch (e) {
      console.error("calendarsGraphSync._syncSingleCalendar error", e);
      throw e;
    }
  }
};