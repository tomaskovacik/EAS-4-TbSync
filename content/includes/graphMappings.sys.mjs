// Mapping helpers: Graph JSON <-> vCard-like object and EAS-style ApplicationData
// Export: graphMappings

export const graphMappings = {
  // Convert Microsoft Graph contact JSON -> vCard-like object
  contactGraphToVCard(graphContact) {
    if (!graphContact) return null;
    const v = {};
    v.fn = graphContact.displayName || `${graphContact.givenName || ""} ${graphContact.surname || ""}`.trim();
    v.n = {
      familyName: graphContact.surname || "",
      givenName: graphContact.givenName || "",
      additionalName: graphContact.middleName || "",
      honorificPrefix: graphContact.title || "",
      honorificSuffix: graphContact.suffix || ""
    };
    v.emails = (graphContact.emailAddresses || []).map(e => ({ address: e.address || "", name: e.name || v.fn || "" }));
    v.telephones = [];
    (graphContact.homePhones || []).forEach(p => v.telephones.push({ type: "home", value: p }));
    (graphContact.businessPhones || []).forEach(p => v.telephones.push({ type: "work", value: p }));
    if (graphContact.mobilePhone) v.telephones.push({ type: "mobile", value: graphContact.mobilePhone });
    v.org = graphContact.companyName || "";
    v.title = graphContact.jobTitle || "";
    v.addresses = [];
    if (graphContact.homeAddress) {
      const a = graphContact.homeAddress;
      v.addresses.push({ type: "home", street: a.street || "", city: a.city || "", region: a.state || a.region || "", postalCode: a.postalCode || "", country: a.countryOrRegion || a.country || "" });
    }
    if (graphContact.businessAddress) {
      const a = graphContact.businessAddress;
      v.addresses.push({ type: "work", street: a.street || "", city: a.city || "", region: a.state || a.region || "", postalCode: a.postalCode || "", country: a.countryOrRegion || a.country || "" });
    }
    if (graphContact.id) v.remoteId = graphContact.id;
    v._graph = graphContact;
    return v;
  },

  // Convert vCard-like object -> Graph contact JSON (best-effort)
  contactVCardToGraph(vcard) {
    if (!vcard) return null;
    const g = {};
    if (vcard.fn) g.displayName = vcard.fn;
    if (vcard.n) { g.givenName = vcard.n.givenName || ""; g.surname = vcard.n.familyName || ""; }
    if (vcard.emails && vcard.emails.length) g.emailAddresses = vcard.emails.map(e => ({ address: e.address, name: e.name || vcard.fn || "" }));
    const homePhones = [], businessPhones = [];
    for (const t of vcard.telephones || []) {
      if (!t || !t.value) continue;
      const typ = (t.type || "").toLowerCase();
      if (typ.includes("home")) homePhones.push(t.value);
      else if (typ.includes("work") || typ.includes("business")) businessPhones.push(t.value);
      else if (typ.includes("mobile")) g.mobilePhone = t.value;
      else businessPhones.push(t.value);
    }
    if (homePhones.length) g.homePhones = homePhones;
    if (businessPhones.length) g.businessPhones = businessPhones;
    if (vcard.org) g.companyName = vcard.org;
    if (vcard.title) g.jobTitle = vcard.title;
    for (const a of vcard.addresses || []) {
      if ((a.type || "").toLowerCase() === "home") {
        g.homeAddress = { street: a.street || "", city: a.city || "", state: a.region || "", postalCode: a.postalCode || "", countryOrRegion: a.country || "" };
      } else {
        g.businessAddress = { street: a.street || "", city: a.city || "", state: a.region || "", postalCode: a.postalCode || "", countryOrRegion: a.country || "" };
      }
    }
    return g;
  },

  // Convert Graph contact (or vCard-like) -> EAS-style ApplicationData object expected by contactsync.setThunderbirdItemFromWbxml
  contactToEasData(graphContactOrVcard) {
    // Accept either Graph contact JSON or vCard-like object (graphMappings.contactGraphToVCard)
    let v = typeof graphContactOrVcard._graph !== "undefined" ? graphContactOrVcard : this.contactGraphToVCard(graphContactOrVcard);
    if (!v) return {};
    const data = {};
    // Basic mapping: use EAS property names used by contactsync.setThunderbirdItemFromWbxml
    data.FileAs = v.fn || "";
    data.FirstName = v.n.givenName || "";
    data.LastName = v.n.familyName || "";
    data.MiddleName = v.n.additionalName || "";
    data.Title = v.n.honorificPrefix || "";
    data.Suffix = v.n.honorificSuffix || "";
    // Email slots (EAS supports Email1Address..3)
    for (let i = 0; i < 3; i++) {
      data[`Email${i+1}Address`] = (v.emails && v.emails[i]) ? v.emails[i].address : "";
    }
    // Phones
    for (const t of v.telephones || []) {
      const typ = (t.type || "").toLowerCase();
      if (typ === "mobile") data.MobilePhoneNumber = t.value;
      else if (typ === "home") data.HomePhoneNumber = t.value;
      else if (typ === "work") data.BusinessPhoneNumber = t.value;
    }
    // Org/title
    data.CompanyName = v.org || "";
    data.JobTitle = v.title || "";
    // Addresses: use first home/work if available
    for (const a of v.addresses || []) {
      if ((a.type || "").toLowerCase() === "home") {
        data.HomeAddressStreet = a.street || "";
        data.HomeAddressCity = a.city || "";
        data.HomeAddressState = a.region || "";
        data.HomeAddressPostalCode = a.postalCode || "";
        data.HomeAddressCountry = a.country || "";
      } else {
        data.BusinessAddressStreet = a.street || "";
        data.BusinessAddressCity = a.city || "";
        data.BusinessAddressState = a.region || "";
        data.BusinessAddressPostalCode = a.postalCode || "";
        data.BusinessAddressCountry = a.country || "";
      }
    }
    // Keep Graph id for mapping
    if (v.remoteId) data._graphId = v.remoteId;
    return data;
  },

  // Convert Graph event -> minimal EAS-like ApplicationData used by calendar write helpers
  eventGraphToEasData(graphEvent) {
    if (!graphEvent) return {};
    const d = {};
    // Map summary/subject
    d.Subject = graphEvent.subject || "";
    // Map body
    if (graphEvent.body && graphEvent.body.content) d.Body = graphEvent.body.content;
    // Map start/end
    if (graphEvent.start) {
      d.StartTime = graphEvent.start.dateTime || "";
      d.StartTimeZone = graphEvent.start.timeZone || "";
    }
    if (graphEvent.end) {
      d.EndTime = graphEvent.end.dateTime || "";
      d.EndTimeZone = graphEvent.end.timeZone || "";
    }
    // All-day
    if (typeof graphEvent.isAllDay !== "undefined") d.IsAllDay = graphEvent.isAllDay ? "1" : "0";
    // Location
    if (graphEvent.location && graphEvent.location.displayName) d.Location = graphEvent.location.displayName;
    // Recurrence (best-effort to EAS representation will be limited)
    if (graphEvent.recurrence) d.Recurrence = graphEvent.recurrence;
    // Remote id
    if (graphEvent.id) d._graphId = graphEvent.id;
    // Attendees: store raw for further processing
    if (graphEvent.attendees) d.Attendees = graphEvent.attendees;
    return d;
  }
};