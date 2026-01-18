// Exports graphMappings
// Small mapping helper between Microsoft Graph contact JSON and a vCard-like object.
// This module intentionally keeps mappings conservative and easy to extend.

export const graphMappings = {
  // Convert Microsoft Graph contact JSON -> vCard-like object
  // vcard is a plain object with properties similar to what contactsync expects:
  // { fn, n: { familyName, givenName }, emails: [{ address, name }], telephones: [{ type, value }], org, title, addresses: [{ type, street, city, region, postalCode, country }] }
  contactGraphToVCard(graphContact) {
    if (!graphContact) return null;
    const v = {};

    // Full name
    v.fn = graphContact.displayName || `${graphContact.givenName || ""} ${graphContact.surname || ""}`.trim();

    // Name components
    v.n = {
      familyName: graphContact.surname || "",
      givenName: graphContact.givenName || "",
      additionalName: graphContact.middleName || "",
      honorificPrefix: graphContact.title || "",
      honorificSuffix: graphContact.suffix || ""
    };

    // Emails
    v.emails = [];
    if (Array.isArray(graphContact.emailAddresses)) {
      for (let e of graphContact.emailAddresses) {
        v.emails.push({ address: e.address || "", name: e.name || v.fn || "" });
      }
    } else if (graphContact.emailAddresses && graphContact.emailAddresses.address) {
      v.emails.push({ address: graphContact.emailAddresses.address, name: graphContact.emailAddresses.name || v.fn || "" });
    }

    // Phones
    v.telephones = [];
    if (graphContact.homePhones && graphContact.homePhones.length) {
      for (const p of graphContact.homePhones) v.telephones.push({ type: "home", value: p });
    }
    if (graphContact.businessPhones && graphContact.businessPhones.length) {
      for (const p of graphContact.businessPhones) v.telephones.push({ type: "work", value: p });
    }
    if (graphContact.mobilePhone) v.telephones.push({ type: "mobile", value: graphContact.mobilePhone });

    // Organization & title
    v.org = graphContact.companyName || "";
    v.title = graphContact.jobTitle || "";

    // Addresses (Graph uses single structured businessAddress/homeAddress)
    v.addresses = [];
    if (graphContact.homeAddress) {
      const a = graphContact.homeAddress;
      v.addresses.push({
        type: "home",
        street: a.street || "",
        city: a.city || "",
        region: a.state || a.region || "",
        postalCode: a.postalCode || "",
        country: a.countryOrRegion || a.country || ""
      });
    }
    if (graphContact.businessAddress) {
      const a = graphContact.businessAddress;
      v.addresses.push({
        type: "work",
        street: a.street || "",
        city: a.city || "",
        region: a.state || a.region || "",
        postalCode: a.postalCode || "",
        country: a.countryOrRegion || a.country || ""
      });
    }

    // Keep original remote id for mapping
    if (graphContact.id) v.remoteId = graphContact.id;

    // Raw Graph object for any additional info
    v._graph = graphContact;

    return v;
  },

  // Convert vCard-like object -> Graph contact JSON (best-effort / minimal)
  contactVCardToGraph(vcard) {
    if (!vcard) return null;
    const g = {};

    if (vcard.fn) g.displayName = vcard.fn;
    if (vcard.n) {
      g.givenName = vcard.n.givenName || "";
      g.surname = vcard.n.familyName || "";
      if (vcard.title) g.title = vcard.title;
    }

    // emailAddresses
    if (vcard.emails && vcard.emails.length) {
      g.emailAddresses = vcard.emails.map(e => ({ address: e.address, name: e.name || vcard.fn || "" }));
    }

    // phones
    const homePhones = [];
    const businessPhones = [];
    for (const t of vcard.telephones || []) {
      if (!t || !t.value) continue;
      const typ = (t.type || "").toLowerCase();
      if (typ.includes("home") || typ === "home") homePhones.push(t.value);
      else if (typ.includes("work") || typ === "work" || typ === "business") businessPhones.push(t.value);
      else if (typ.includes("mobile") || typ === "mobile") g.mobilePhone = t.value;
      else businessPhones.push(t.value);
    }
    if (homePhones.length) g.homePhones = homePhones;
    if (businessPhones.length) g.businessPhones = businessPhones;

    // organization/title
    if (vcard.org) g.companyName = vcard.org;
    if (vcard.title) g.jobTitle = vcard.title;

    // addresses - map first home/work found
    for (const a of vcard.addresses || []) {
      if ((a.type || "").toLowerCase() === "home") {
        g.homeAddress = {
          street: a.street || "",
          city: a.city || "",
          state: a.region || "",
          postalCode: a.postalCode || "",
          countryOrRegion: a.country || ""
        };
      } else {
        // treat as business
        g.businessAddress = {
          street: a.street || "",
          city: a.city || "",
          state: a.region || "",
          postalCode: a.postalCode || "",
          countryOrRegion: a.country || ""
        };
      }
    }

    return g;
  }
};