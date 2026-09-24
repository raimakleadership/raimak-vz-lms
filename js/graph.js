// Raimak LMS - Graph API / SharePoint Data Layer v3.0

const Graph = (() => {
  const base = Config.sharePoint.graphBase;
  const host = Config.sharePoint.hostname;
  const lists = Config.sharePoint.lists;
  let siteIds = { leadship: null, team: null };
  let agentCache = null;

  // ============================================================
  // 📱 SMS HISTORY & MESSAGING API (Delta Cached)
  // ============================================================
  const SMSApi = {
    /**
     * Delta Syncs the SMS History from SharePoint, caching it in IndexedDB and RAM.
     */
    syncSMSHistory: async function (
      lastSyncDate = null,
      existingMessages = [],
    ) {
      await resolveSiteIds();

      if (!existingMessages || existingMessages.length === 0) {
        if (typeof LocalDB !== "undefined" && LocalDB.getAllItems) {
          existingMessages = await LocalDB.getAllItems("sms_history");
          existingMessages.sort(
            (a, b) => new Date(a.timestamp) - new Date(b.timestamp),
          );
        }
      }

      let effectiveSyncDate =
        existingMessages.length === 0 ? null : lastSyncDate;
      const listId = Config.sharePoint.lists.smsHistory;

      // 🚀 THE FIX 1: Add "Modified" to the $select string
      let url =
        base +
        "/sites/" +
        siteIds.team +
        "/lists/" +
        listId +
        "/items?$expand=fields($select=Title,Direction,MessageBody,Agent,DeliveryStatus,Created,Modified)&$top=5000";

      if (effectiveSyncDate) {
        const safeDate = effectiveSyncDate.split(".")[0] + "Z";
        // 🚀 THE FIX 2: Ask SharePoint for recently *Modified* items, not just recently Created!
        url += `&$filter=fields/Modified gt '${safeDate}'`;
      }

      const raw = await getAllItems(url);

      if (raw.length === 0) {
        return { updatedMessages: existingMessages, newSyncDate: lastSyncDate };
      }

      const newMessages = raw.map((item) => ({
        id: item.id,
        cbr: item.fields.Title,
        direction: item.fields.Direction,
        message: item.fields.MessageBody,
        agent: item.fields.Agent,
        status: item.fields.DeliveryStatus,
        timestamp: item.fields.Created,
        modified:
          item.fields.Modified ||
          item.lastModifiedDateTime ||
          item.fields.Created, // Track edits!
      }));

      const msgMap = new Map();
      existingMessages.forEach((m) => msgMap.set(m.id, m));
      newMessages.forEach((m) => msgMap.set(m.id, m));

      const finalizedMessages = Array.from(msgMap.values());
      finalizedMessages.sort(
        (a, b) => new Date(a.timestamp) - new Date(b.timestamp),
      );

      if (typeof LocalDB !== "undefined" && LocalDB.saveItems) {
        await LocalDB.saveItems("sms_history", newMessages);
      }

      // 🚀 THE FIX 3: Update the sync tracker timestamp using the new Modified date
      const validTimestamps = newMessages
        .map((m) => new Date(m.modified).getTime())
        .filter((t) => !isNaN(t));

      let newLastSyncDate = lastSyncDate;
      if (validTimestamps.length > 0) {
        const maxTime = Math.max(...validTimestamps);
        newLastSyncDate = new Date(maxTime).toISOString();
        localStorage.setItem("RaimakVZSMSLastSyncDate", newLastSyncDate);
      }

      return {
        updatedMessages: finalizedMessages,
        newSyncDate: newLastSyncDate,
      };
    },

    /**
     * Reads strictly from local RAM. Zero Graph API calls. (Includes Omni-Matcher)
     */
    getLeadMessages: async function (rawCbr) {
      if (!rawCbr) return [];

      const cleanPhone = String(rawCbr).replace(/\D/g, "");
      const base10 = cleanPhone.slice(-10);
      const plus1 = "+1" + base10;
      const just1 = "1" + base10;
      const dashed = `${base10.slice(0, 3)}-${base10.slice(3, 6)}-${base10.slice(6)}`;

      const allowedFormats = [base10, plus1, just1, dashed];

      // 🚀 THE FIX: Dropped "window." so it accesses your const State properly
      const messages = (State.smsHistory || []).filter((m) =>
        allowedFormats.includes(m.cbr),
      );

      // Ensure chronological order just in case
      return messages.sort(
        (a, b) => new Date(a.timestamp) - new Date(b.timestamp),
      );
    },

    /**
     * Drops a new message into SharePoint and optimistically updates local RAM.
     */
    sendTextMessage: async function (cbr, messageBody) {
      await resolveSiteIds();

      const listId = Config.sharePoint.lists.smsHistory;
      const url =
        base + "/sites/" + siteIds.team + "/lists/" + listId + "/items";

      const agentIdentifier =
        State.currentUser && State.currentUser.email
          ? State.currentUser.email
          : "System";

      const fields = {
        Title: cbr,
        Direction: "Outbound",
        MessageBody: messageBody,
        Agent: agentIdentifier,
        DeliveryStatus: "Pending",
      };

      try {
        const response = await apiFetch(url, {
          method: "POST",
          body: { fields },
        });

        const newMsg = {
          id: response.id,
          cbr: response.fields.Title,
          direction: response.fields.Direction,
          message: response.fields.MessageBody,
          agent: response.fields.Agent,
          status: response.fields.DeliveryStatus,
          timestamp: response.fields.Created || new Date().toISOString(),
        };

        // 🚀 THE FIX: Dropped "window." here as well!
        if (typeof State !== "undefined" && State.smsHistory) {
          State.smsHistory.push(newMsg);
        }

        if (typeof LocalDB !== "undefined" && LocalDB.saveItems) {
          LocalDB.saveItems("sms_history", [newMsg]);
        }

        return newMsg;
      } catch (error) {
        console.error("❌ Failed to queue outbound SMS:", error);
        throw error;
      }
    },
  };

  // ── Generic Fetch ──────────────────────────────────────────
  async function apiFetch(url, options = {}) {
    // 1. Auth Check
    const token = await Auth.getToken();
    if (!token) {
      console.warn("No auth token available — redirecting to sign in.");
      Auth.signIn();
      return null;
    }

    // 2. Setup Default Headers & Merge Custom Ones
    const headers = {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
      // 🚀 THE PREFERENCE HEADER: Tells SharePoint to try harder on large lists
      // Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
      ...options.headers, // This allows updateLead to inject "If-Match": "*"
    };

    // 3. Prepare Fetch Options
    const method = options.method || "GET";
    const fetchOpts = {
      method: method,
      headers: headers,
    };

    // If a body was passed, ensure it is stringified
    if (options.body) {
      fetchOpts.body =
        typeof options.body === "string"
          ? options.body
          : JSON.stringify(options.body);
    }

    const maxRetries = options.maxRetries || 3;

    // 4. Execution Loop with Retry Logic
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      const res = await fetch(url, fetchOpts);

      if (res.ok) {
        if (res.status === 204) return null;
        return res.json();
      }

      // Intercept 429 Throttling
      if (res.status === 429) {
        const retryAfterStr = res.headers.get("Retry-After");
        const waitMs = retryAfterStr
          ? parseInt(retryAfterStr) * 1000
          : Math.pow(2, attempt) * 1000;

        console.warn(
          `Graph API Throttled! Pausing for ${waitMs}ms (Attempt ${attempt} of ${maxRetries})`,
        );
        await new Promise((resolve) => setTimeout(resolve, waitMs));
        continue;
      }

      // Handle other errors
      const err = await res.json().catch(() => ({}));
      throw new Error((err.error && err.error.message) || "HTTP " + res.status);
    }

    throw new Error("Microsoft Graph request failed after max retries.");
  }

  // ── Resolve Site IDs ───────────────────────────────────────
  async function resolveSiteIds() {
    if (siteIds.leadship && siteIds.team) return;

    try {
      const [s1, s2] = await Promise.all([
        // 🚀 Pattern Shift: Passing an empty object or { method: "GET" }
        // ensures it hits the upgraded apiFetch signature correctly.
        apiFetch(
          base + "/sites/" + host + ":/" + Config.sharePoint.sites.leadship,
          { method: "GET" },
        ),
        apiFetch(
          base + "/sites/" + host + ":/" + Config.sharePoint.sites.team,
          { method: "GET" },
        ),
      ]);

      siteIds.leadship = s1.id;
      siteIds.team = s2.id;
    } catch (err) {
      console.error(
        "Critical Error: Could not resolve SharePoint Site IDs.",
        err,
      );
      throw err;
    }
  }

  // ── Build Agent ID Cache ───────────────────────────────────
  async function resolveAgentCache() {
    if (agentCache) return agentCache;
    await resolveSiteIds();
    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.contractorList +
      "/items?expand=fields&$top=500";
    const raw = await getAllItems(url);
    agentCache = {};
    raw.forEach(function (item) {
      const name = (
        item.fields &&
        (item.fields.Title || item.fields.ContractorName || "")
      )
        .toLowerCase()
        .trim();
      if (name) agentCache[name] = parseInt(item.id, 10);
    });
    return agentCache;
  }

  async function resolveAgentId(agentName) {
    if (!agentName) return null;
    const cache = await resolveAgentCache();
    return cache[agentName.toLowerCase().trim()] || null;
  }

  async function assignAgent(itemId, agentName) {
    await updateLead(itemId, { Agent_x0020_Assigned: agentName });
  }

  // ── Paginate ───────────────────────────────────────────────
  async function getAllItems(url, customOptions = {}) {
    const items = [];
    let next = url;

    try {
      while (next) {
        const data = await apiFetch(next, {
          method: "GET",
          maxRetries: customOptions.maxRetries || 5, // 🚀 Default to 5 retries instead of 3
          headers: {
            Prefer: "allow-throttleable-queries",
            ...customOptions.headers,
          },
        });

        if (data && data.value) {
          items.push(...data.value);
        }

        next = data["@odata.nextLink"] || null;
      }
    } catch (error) {
      console.error("❌ Pagination failed on URL:", next, error.message);
      if (window.UI && UI.showToast) {
        UI.showToast("Network interrupted. Partial data loaded.", "warning");
      }
    }

    return items;
  }

  // ── Self-Injecting & Hoisted Streamed Pager for Cold Boots ──
  async function getAllItemsStreamed(storeName, initialUrl) {
    const items = [];
    let next = initialUrl;
    let pageCount = 0;

    // 🚀 1. FIND OR BUILD THE MODAL WITH MAX Z-INDEX
    let modal = document.getElementById("cold-boot-overlay");
    if (!modal) {
      modal = document.createElement("div");
      modal.id = "cold-boot-overlay";
      modal.style.cssText =
        "position: fixed; inset: 0; background: rgba(0, 0, 0, 0.92); backdrop-filter: blur(8px); z-index: 2147483647; display: flex; align-items: center; justify-content: center; padding: 20px;";
      modal.innerHTML = `
        <div class="card" style="max-width: 440px; width: 100%; padding: 32px; text-align: center; border: 1px solid #3a181b; background: #121216; box-shadow: 0 0 50px rgba(237, 28, 36, 0.25); border-radius: 16px;">
          <div style="font-size: 48px; margin-bottom: 16px;">⚡</div>
          <h2 style="font-family: var(--font-head, sans-serif); font-size: 24px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px; color: #f8fafc;">
            Caching Verizon Floor Data
          </h2>
          <p style="color: #94a3b8; font-size: 13px; margin-bottom: 24px; line-height: 1.5;">
            Downloading active leads to your device for instant offline search and zero-lag dialing. This only runs on first login.
          </p>
          <div style="width: 100%; background: #0a0a0c; height: 10px; border-radius: 5px; overflow: hidden; border: 1px solid #22222a; margin-bottom: 12px;">
            <div id="cold-boot-fill" style="width: 0%; height: 100%; background: linear-gradient(90deg, #cd040b, #ed1c24); transition: width 0.3s ease; box-shadow: 0 0 10px rgba(237, 28, 36, 0.6);"></div>
          </div>
          <div style="display: flex; justify-content: space-between; font-family: var(--font-mono, monospace); font-size: 11px; color: #94a3b8;">
            <span id="cold-boot-status">Connecting to Microsoft Graph...</span>
            <span id="cold-boot-percent">0%</span>
          </div>
        </div>
      `;
    }

    // 🚀 2. THE HOISTING FIX: Always move it to the very end of document.body
    // so it renders over top of #app-shell even after showAppShell() runs!
    document.body.appendChild(modal);
    modal.style.display = "flex";

    const fill = document.getElementById("cold-boot-fill");
    const statusText = document.getElementById("cold-boot-status");
    const percentText = document.getElementById("cold-boot-percent");

    try {
      while (next) {
        pageCount++;
        if (statusText)
          statusText.textContent = `Downloading floor batch #${pageCount}...`;

        const data = await apiFetch(next, {
          method: "GET",
          maxRetries: 7,
          headers: { Prefer: "allow-throttleable-queries" },
        });

        if (data && data.value) {
          items.push(...data.value);

          // 3. Stream each page directly into IndexedDB
          const normalizedChunk = data.value.map(normalizeLeadItem);
          await LocalDB.saveItems(storeName, normalizedChunk);

          // 4. Update Progress Bar (based on ~30,000 total floor leads)
          const estimatedTotal = 30000;
          const currentPct = Math.min(
            98,
            Math.round((items.length / estimatedTotal) * 100),
          );
          if (fill) fill.style.width = `${currentPct}%`;
          if (percentText) percentText.textContent = `${currentPct}%`;
          if (statusText)
            statusText.textContent = `Cached ${items.length.toLocaleString()} leads...`;
        }

        next = data["@odata.nextLink"] || null;

        // ⏱️ 5. Polite 350ms TCPA/Graph pacing between pages
        if (next) {
          await new Promise((resolve) => setTimeout(resolve, 350));
        }
      }

      if (fill) fill.style.width = "100%";
      if (percentText) percentText.textContent = "100%";
      if (statusText) statusText.textContent = "Floor cache ready!";
      await new Promise((resolve) => setTimeout(resolve, 600));
    } catch (error) {
      console.error("❌ Cold boot streaming failed:", error);
      if (window.UI && UI.showToast) {
        UI.showToast(
          "Sync interrupted. Partially cached floor data.",
          "warning",
        );
      }
    } finally {
      if (modal) modal.style.display = "none";
    }

    return items;
  }
  // ============================================================
  //  LEADS
  // ============================================================

  async function getLeads(lastSyncDate = null, existingLeads = []) {
    await resolveSiteIds();

    // 🚀 STEP 1: Load from IndexedDB
    if (!existingLeads || existingLeads.length === 0) {
      existingLeads = await LocalDB.getAllItems("leads");
      console.log(
        `📦 [Graph.getLeads] Loaded ${existingLeads.length} leads from IndexedDB.`,
      );
    }

    // 🚀 STEP 2: GUARANTEED COLD BOOT TRIGGER
    // If IndexedDB has fewer than 1,000 leads, force Cold Boot mode!
    const isColdBoot = existingLeads.length < 1000;

    const expandQuery = "expand=fields";
    let url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.leadsList +
      `/items?${expandQuery}&$select=id,lastModifiedDateTime,createdDateTime&$top=5000`;

    // Only apply delta filter if we are NOT cold booting
    if (!isColdBoot && lastSyncDate && typeof lastSyncDate === "string") {
      const safeDate = lastSyncDate.split(".")[0] + "Z";
      url += `&$filter=fields/Modified gt '${safeDate}'`;
    }

    // 🚀 STEP 3: Fetch with Streamed Progress Modal if Cold Booting
    let raw = [];
    if (isColdBoot) {
      raw = await getAllItemsStreamed("leads", url);
    } else {
      raw = await getAllItems(url);
    }

    if (raw.length === 0) {
      return existingLeads.filter(
        (l) =>
          l.status !== "D2D Lead" &&
          l.status !== "TDM Non-Reg" &&
          l.status !== "Deleted",
      );
    }

    const updatedBatch = raw.map(normalizeLeadItem);

    // Save to IndexedDB if we didn't already stream it page-by-page
    if (!isColdBoot) {
      await LocalDB.saveItems("leads", updatedBatch);
    }

    const leadMap = new Map();
    existingLeads.forEach((l) => leadMap.set(l.id, l));
    updatedBatch.forEach((l) => leadMap.set(l.id, l));

    const finalizedLeads = Array.from(leadMap.values()).filter(
      (l) =>
        l.status !== "D2D Lead" &&
        l.status !== "TDM Non-Reg" &&
        l.status !== "Deleted",
    );

    const validTimestamps = updatedBatch
      .map((l) => new Date(l.modified || l.lastModifiedDateTime).getTime())
      .filter((t) => !isNaN(t));

    if (validTimestamps.length > 0) {
      const maxTime = Math.max(...validTimestamps);
      localStorage.setItem(
        "RaimakVZLeadsLastSyncDate",
        new Date(maxTime).toISOString(),
      );
    }

    return finalizedLeads;
  }

  // ── Self-Injecting Streamed Pager for Cold Boots ──

  async function getNextLeadForAgent(agentEmail) {
    await resolveSiteIds();
    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.leadsList +
      "/items?expand=fields&$top=500";
    const raw = await getAllItems(url);
    const leads = raw.map(normalizeLeadItem);
    return (
      leads.find(
        (l) => l.status === "New" && !l.assignedTo && !isInCoolOff(l),
      ) || null
    );
  }

  async function addLead(fields) {
    await resolveSiteIds();
    const url =
      base + "/sites/" + siteIds.team + "/lists/" + lists.leadsList + "/items";

    const res = await apiFetch(url, {
      method: "POST",
      body: { fields },
    });

    return normalizeLeadItem(res);
  }

  async function updateLead(itemId, fields) {
    await resolveSiteIds();

    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.leadsList +
      "/items/" +
      itemId +
      "/fields";

    const options = {
      method: "PATCH",
      body: JSON.stringify(fields),
      headers: {
        "Content-Type": "application/json",
        "If-Match": "*",
      },
    };

    try {
      // Try the normal update first
      return await apiFetch(url, options);
    } catch (err) {
      const errMsg = (err.message || "").toLowerCase();

      // 👻 THE GHOST LEAD INTERCEPTOR
      if (errMsg.includes("404") || errMsg.includes("not found")) {
        console.warn(
          `Lead ${itemId} missing from server! Checking local memory...`,
        );

        // 1. Find the INDEX of the ghost using type-safe String coercion
        const deadLeadIndex =
          window.State && window.State.leads
            ? window.State.leads.findIndex(
                (l) => String(l.id) === String(itemId),
              )
            : -1;

        if (deadLeadIndex === -1) {
          throw new Error(
            "Lead missing from server and no local memory found to rebuild it.",
          );
        }

        const deadLead = window.State.leads[deadLeadIndex];

        // 🚀 THE PURGE MECHANIC: Strictly enforces Name & Address
        if (!deadLead.name || !deadLead.address) {
          console.warn(
            `Lead ${itemId} lacks mandatory Name/Address. Purging from local cache.`,
          );
          window.State.leads.splice(deadLeadIndex, 1);
          throw new Error(
            "Lead missing from server and lacked minimum data to resurrect. It has been deleted locally.",
          );
        }

        console.warn(`Resurrecting Lead ${itemId}...`);

        // 2. REVERSE MAP: Building the payload using your normalizeLeadItem schema
        const rebuiltPayload = {
          // THE MANDATORIES
          Title: deadLead.name,
          Address: deadLead.address,

          // THE NICE-TO-HAVES (Pulled from local cache, ignoring Phone/CBR/BTN)
          FirstName: deadLead.firstName || "",
          LastName: deadLead.lastName || "",
          Email: deadLead.email || "",
          Status: deadLead.status || "New",
          AssignedTo: deadLead.assignedTo || "",
          City: deadLead.city || "",
          State: deadLead.state || "",
          Zip: deadLead.zip || "",
          Notes: deadLead.notes || "",
          LeadType: deadLead.leadType || "",
          PreviousAgents: deadLead.previousAgents || "",

          // 3. Drop the brand new edits right on top of the resurrected data
          ...fields,
        };

        // 4. Fire your existing addLead function to create the new row in SharePoint
        const resurrectedLead = await addLead(rebuiltPayload);

        // 5. Swap the ID in local RAM so the UI knows about the newly generated SharePoint ID
        deadLead.id = resurrectedLead.id;

        // Pretend everything went perfectly
        return resurrectedLead;
      }

      // If it failed for any other reason (like a 500 error), throw it normally
      throw err;
    }
  }

  /*  async function deleteLead(itemId) {
    await resolveSiteIds();
    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.leadsList +
      "/items/" +
      itemId;

    // 🚀 Pattern Shift: Using the options object with If-Match to prevent version conflicts
    await apiFetch(url, {
      method: "DELETE",
      headers: {
        "If-Match": "*",
      },
    });
 } */

  function normalizeLeadItem(item) {
    const f = item.fields || {};
    const first = f.FirstName || f.First_x0020_Name || "";
    const last = f.LastName || f.Last_x0020_Name || "";
    const name = (first + " " + last).trim() || f.Title || f.LeadName || "";
    return {
      id: item.id,
      name: name,
      firstName: first,
      lastName: last,
      email: f.Email || f.EmailAddress || "",
      phone: f.Phone || f.PhoneNumber || "",
      status: f.Status || "New",
      source: f.Campaign || f.LeadSource || f.Source || "",
      assignedTo:
        f.Agent_x0020_Assigned ||
        f.AgentAssigned ||
        f.AssignedTo ||
        f.Agent ||
        "",
      notes: f.Notes || "",
      address: f.WorkAddress || f.Address || "",
      city: f.WorkCity || f.City || "",
      state: f.State || "",
      zip: f.Zip || f.ZipCode || "",
      cbr: f.CBR || "",
      btn: f.BTN || "",
      lockFlag: f.LockFlag || false,
      callbackAt: f.CallbackDateTime || null,
      lastContacted: f.LastTouchedOn || f.LastContacted || null,
      createdAt: item.createdDateTime || f.Created || null,
      modified: item.lastModifiedDateTime || null,
      leadType:
        f.Lead_x0020_Type || f.Type || f.Item_x0020_Type || f.LeadType || "",
      currentMRC:
        f.MonthlyRecurringCharge_x0028_MRC || f.CurrentMRC || f.MRC || "",

      // 🚀 NEW VERIZON FIELDS
      currentProvider: f.CurrentProvider || "",
      numberOfLines: f.NumberOfLines || "",
      phoneStatus: f.PhoneStatus || "",
      quotedAt: f.QuotedAt || "",

      previousAgents: f.PreviousAgents || "",
      // 💬 VERIZON & SMS FIELDS
      leadId: f.LeadID || "",
      SMSOptOut: f.SMSOptOut === true || f.SMSOptOut === "true",
      SMSLastContacted: f.SMSLastContacted || null,
      SMSUnread: f.SMSUnread === true || f.SMSUnread === "true",
    };
  }

  // ============================================================
  //  CONTRACTORS / AGENTS
  // ============================================================

  async function getContractors() {
    await resolveSiteIds();
    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.contractorList +
      "/items?expand=fields&$top=500";
    const raw = await getAllItems(url);
    return raw.map((item) => {
      const f = item.fields || {};
      return {
        id: item.id,
        name: f.Title || f.ContractorName || "",
        email: f.Email || "",
        phone: f.Phone || "",
        role: f.Role || "Agent",
        active: f.Active !== undefined ? f.Active : true,
      };
    });
  }

  async function getAgentScores() {
    await resolveSiteIds();

    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.agentScores + // <-- Now using the dynamic config ID
      "/items?expand=fields($select=AgentEmail,AgentName,CurrentPoints,LifetimePoints)&$top=500";

    const raw = await getAllItems(url);

    return raw.map((item) => {
      const f = item.fields || {};
      return {
        id: item.id,
        AgentEmail: f.AgentEmail || "",
        AgentName: f.AgentName || "",
        CurrentPoints:
          typeof f.CurrentPoints === "number" ? f.CurrentPoints : 0,
        LifetimePoints:
          typeof f.LifetimePoints === "number" ? f.LifetimePoints : 0,
      };
    });
  }

  async function createAgentScore(email, name) {
    await resolveSiteIds();

    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.agentScores +
      "/items";

    const payload = {
      fields: {
        AgentEmail: email,
        AgentName: name,
        CurrentPoints: 0,
        LifetimePoints: 0,
      },
    };

    // 🚀 Pattern Shift: Using the options object for the POST request
    const res = await apiFetch(url, {
      method: "POST",
      body: payload,
    });

    return {
      id: res.id,
      AgentEmail: email,
      AgentName: name,
      CurrentPoints: 0,
      LifetimePoints: 0,
    };
  }

  async function checkLedgerForDuplicate(leadId, actionType) {
    if (!leadId || !actionType) return false;

    await resolveSiteIds();

    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.agentScoresLedger +
      `/items?$expand=fields&$filter=fields/LeadID eq '${leadId}' and fields/ActionType eq '${actionType}'`;

    try {
      // 🚀 Pattern Shift: Using the options object for the GET request
      const res = await apiFetch(url, { method: "GET" });

      if (res && res.value && res.value.length > 0) {
        return true;
      }
      return false;
    } catch (err) {
      console.error("Ledger Check Error:", err);
      // Safe-fail: Assume duplicate if check fails to prevent double-point awarding
      return true;
    }
  }

  // ── THE RECEIPT: Write the transaction to the Ledger ──
  async function writeLedgerTransaction(
    agentEmail,
    actionType,
    pointValue,
    leadId = "",
  ) {
    await resolveSiteIds();

    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.agentScoresLedger +
      "/items";

    // Generate a unique transaction ID
    const transactionId = `${agentEmail}_${actionType}_${Date.now()}`;

    const payload = {
      fields: {
        Title: transactionId,
        AgentEmail: agentEmail,
        ActionType: actionType,
        PointValue: pointValue,
        LeadID: leadId,
      },
    };

    // 🚀 Pattern Shift: Using the options object for the POST request
    await apiFetch(url, {
      method: "POST",
      body: payload,
    });
  }

  async function updateAgentScore(itemId, currentPoints, lifetimePoints) {
    await resolveSiteIds();

    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.agentScores +
      "/items/" +
      itemId +
      "/fields";

    const payload = {
      CurrentPoints: currentPoints,
      LifetimePoints: lifetimePoints,
    };

    // 🚀 Pattern Shift: Using the options object with If-Match to prevent score conflicts
    await apiFetch(url, {
      method: "PATCH",
      body: payload,
      headers: {
        "If-Match": "*",
      },
    });
  }

  // ============================================================
  //  ACTIVITY LOG
  // ============================================================
  async function getActivityLog(
    lastSyncDate = null,
    existingLogs = [],
    isDeltaRefresh = false,
  ) {
    await resolveSiteIds();

    // 🚀 STEP 1: If RAM is empty (F5 refresh), try to load from the Local Database first
    if (!existingLogs || existingLogs.length === 0) {
      existingLogs = await LocalDB.getAllItems("activity_logs");

      // Sort them newest first so the UI stays consistent
      existingLogs.sort(
        (a, b) => new Date(b.timestamp) - new Date(a.timestamp),
      );

      if (existingLogs.length > 0) {
        UI.showToast(
          `🚀 Loaded ${existingLogs.length} logs from local cache.`,
          "success",
        );
      }
    }

    // 🚀 STEP 2: Decide if we actually need to download anything
    // If RAM was empty and the Database was empty, we do a full sync (null)
    // Otherwise, we use the saved lastSyncDate to get the delta
    let effectiveSyncDate = existingLogs.length === 0 ? null : lastSyncDate;

    const selectedFields =
      "LeadID,LeadId,Title,LeadName,ActionType,Action,Activity,AgentEmail,Agent,Notes,Created";
    let url =
      base +
      "/sites/" +
      siteIds.leadship +
      "/lists/" +
      lists.activityLog +
      `/items?expand=fields($select=${selectedFields})&$select=id,createdDateTime&$top=5000`;

    if (effectiveSyncDate) {
      // The millisecond-scrubbing fix for Microsoft Graph
      const safeDate = effectiveSyncDate.split(".")[0] + "Z";
      url += `&$filter=fields/Created gt '${safeDate}'`;
    }

    const raw = await getAllItems(url);

    // If no new items exist, we're done! Return what we have.
    if (raw.length === 0) {
      if (isDeltaRefresh) UI.showToast("✅ Logs are up to date.", "success");
      return { updatedLogs: existingLogs, newSyncDate: lastSyncDate };
    }

    // 🚀 STEP 3: Map the new items from Microsoft
    const newLogs = raw.map((item) => {
      const f = item.fields || {};
      return {
        id: item.id, // Primary Key for IndexedDB
        leadId: String(f.LeadID || f.LeadId || ""),
        leadName: f.Title || f.LeadName || "",
        action: f.ActionType || f.Action || f.Activity || "",
        agent: f.AgentEmail || f.Agent || "",
        agentEmail: f.AgentEmail || "",
        notes: f.Notes || "",
        timestamp: item.createdDateTime || f.Created || null,
      };
    });

    // 🚀 STEP 4: Save the new items to IndexedDB permanently!
    // This happens in the background so the UI doesn't lag.
    await LocalDB.saveItems("activity_logs", newLogs);

    // 🚀 STEP 5: Merge and Sort
    // New items go to the front, followed by the existing local logs
    const finalizedLogs = [...newLogs.reverse(), ...existingLogs];

    // 🚀 STEP 6: Update the Sync Date (High-Water Mark)
    const validTimestamps = newLogs
      .map((log) => new Date(log.timestamp).getTime())
      .filter((time) => !isNaN(time));

    let newLastSyncDate = lastSyncDate;
    if (validTimestamps.length > 0) {
      const maxTime = Math.max(...validTimestamps);
      newLastSyncDate = new Date(maxTime).toISOString();
      localStorage.setItem("RaimakVZActivityLastSyncDate", newLastSyncDate);
    }

    if (effectiveSyncDate && isDeltaRefresh) {
      UI.showToast(`✅ Synced ${newLogs.length} new logs.`, "success");
    }

    return {
      updatedLogs: finalizedLogs,
      newSyncDate: newLastSyncDate,
    };
  }

  async function getActivityLogForToday() {
    await resolveSiteIds();

    const selectedFields =
      "LeadID,LeadId,Title,LeadName,ActionType,Action,Activity,AgentEmail,Agent,Notes,Created";

    const url =
      base +
      "/sites/" +
      siteIds.leadship +
      "/lists/" +
      lists.activityLog +
      `/items?expand=fields($select=${selectedFields})&$select=id,createdDateTime&$top=5000`;

    let next = url;
    const todayStr = new Date().toDateString();
    const todayLogs = [];

    try {
      while (next) {
        // 🚀 Pattern Shift: Method and Headers wrapped into the options object
        const response = await apiFetch(next, {
          method: "GET",
          headers: {
            Prefer: "allow-throttleable-queries",
          },
        });

        const items = response.value || [];

        for (const item of items) {
          const f = item.fields || {};
          const timestamp = item.createdDateTime || f.Created;

          if (!timestamp) continue;

          if (new Date(timestamp).toDateString() === todayStr) {
            todayLogs.push({
              id: item.id,
              leadId: String(f.LeadID || f.LeadId || ""),
              leadName: f.Title || f.LeadName || "",
              action: f.ActionType || f.Action || f.Activity || "",
              agent: f.AgentEmail || f.Agent || "",
              agentEmail: f.AgentEmail || "",
              notes: f.Notes || "",
              timestamp: timestamp,
            });
          }
        }

        next = response["@odata.nextLink"] || null;
      }
    } catch (err) {
      console.error("Failed to fetch today's logs:", err);
    }

    // Newest logs at index 0 for the UI
    return todayLogs.reverse();
  }

  async function logActivity(entry) {
    await resolveSiteIds();
    const url =
      base +
      "/sites/" +
      siteIds.leadship + // 🚀 Using your 'leadship' site ID
      "/lists/" +
      lists.activityLog +
      "/items";

    // 🚀 THE FIX: Wrap the method and body into a single options object
    // to match your upgraded apiFetch signature.
    await apiFetch(url, {
      method: "POST",
      body: {
        fields: entry,
      },
    });
  }

  // Get today's sold leads based on activity log entries.
  // Fetches leads directly to avoid timing issues with State.leads.
  function getTodaySales(todayLogs = []) {
    // 1. THE MAPPER: Build the translation dictionary from contractors
    const nameLookup = {};
    (State.contractors || []).forEach((c) => {
      if (c.email) nameLookup[c.email.toLowerCase().trim()] = c.name;
    });

    const sales = [];
    const seenLeadIds = new Set(); // The Bouncer
    const todayStr = new Date().toDateString(); // Get today's date for filtering

    todayLogs.forEach(function (e) {
      // 🛑 CRITICAL: Make sure the log actually happened today!
      const isToday =
        e.timestamp && new Date(e.timestamp).toDateString() === todayStr;

      if (isToday && e.action === "Status: " + Config.soldStatus && e.leadId) {
        // Only count the sale if we haven't seen this exact lead ID yet today
        if (!seenLeadIds.has(e.leadId)) {
          seenLeadIds.add(e.leadId);

          // 2. THE INTERCEPT: Translate the raw agent ID before saving it
          const rawAgent = e.agent || "Unknown";
          const displayName =
            nameLookup[rawAgent.toLowerCase().trim()] || rawAgent;

          sales.push({
            id: e.leadId,
            name: e.leadName || "Unknown Lead",
            soldBy: displayName, // Saves "John Doe" instead of "jdoe@..."
            assignedTo: displayName, // Saves "John Doe" instead of "jdoe@..."
            modified: e.timestamp,
            saleTime: e.timestamp,
          });
        }
      }
    });

    return sales;
  }

  // Get daily activity stats per agent for the report.
  // Maps agent emails back to display names via the contractors list.
  async function getDailyStats() {
    const log = State.activityLog || [];
    const today = new Date().toDateString();

    const todayEntries = log.filter(
      (e) => e.timestamp && new Date(e.timestamp).toDateString() === today,
    );

    const emailToName = {};
    (State.contractors || []).forEach((c) => {
      if (c.email) emailToName[c.email.toLowerCase().trim()] = c.name;
    });

    const stats = {};
    for (const entry of todayEntries) {
      const agentEmail = (entry.agent || "").toLowerCase().trim();
      const agent = emailToName[agentEmail] || entry.agent || "Unknown";

      if (!stats[agent]) {
        stats[agent] = {
          agent,
          actions: [],
          uniqueLeads: new Set(),
          uniqueSales: new Set(),
        };
      }

      const isContact =
        entry.action &&
        (entry.action.indexOf("Status:") === 0 ||
          entry.action.includes("Contact"));
      if (isContact && entry.leadId) stats[agent].uniqueLeads.add(entry.leadId);
      if (entry.action === "Status: " + Config.soldStatus && entry.leadId)
        stats[agent].uniqueSales.add(entry.leadId);

      stats[agent].actions.push(entry);
    }

    return Object.values(stats)
      .map(function (s) {
        // 🚀 CALCULATE AVERAGE CADENCE
        // Sort actions by time (oldest to newest)
        const sorted = s.actions.sort(
          (a, b) => new Date(a.timestamp) - new Date(b.timestamp),
        );
        let totalDiff = 0;
        let intervalCount = 0;

        for (let i = 1; i < sorted.length; i++) {
          const diff =
            new Date(sorted[i].timestamp) - new Date(sorted[i - 1].timestamp);

          // Only count intervals under 30 minutes (exclude lunch/long breaks)
          if (diff > 0 && diff < 30 * 60 * 1000) {
            totalDiff += diff;
            intervalCount++;
          }
        }

        const avgMs = intervalCount > 0 ? totalDiff / intervalCount : 0;
        const mins = Math.floor(avgMs / 60000);
        const secs = Math.floor((avgMs % 60000) / 1000);

        s.avgTime = avgMs > 0 ? `${mins}m ${secs}s` : "—";
        s.contacts = s.uniqueLeads.size;
        s.sold = s.uniqueSales.size;

        delete s.uniqueLeads;
        delete s.uniqueSales;
        return s;
      })
      .sort((a, b) => b.contacts - a.contacts);
  }

  // ============================================================
  //  BUSINESS RULES
  // ============================================================

  function applyBusinessRules(leads, contractors) {
    const now = new Date();
    const { coolOffDays, maxLeadsPerAgent, recycleAfterDays } = Config.rules;

    const agentCounts = {};
    for (const lead of leads) {
      if (lead.assignedTo && !Config.terminalStatuses.includes(lead.status)) {
        agentCounts[lead.assignedTo] = (agentCounts[lead.assignedTo] || 0) + 1;
      }
    }

    return leads.map(function (lead) {
      const flags = [];

      // 🚀 THE BOUNCER: Is an agent actually holding this lead?
      // If it's unassigned (in the free pool), it cannot be "recycled".
      const isAssigned = !!(lead.assignedTo && lead.assignedTo.trim() !== "");

      if (lead.lastContacted) {
        const daysSince = (now - new Date(lead.lastContacted)) / 86400000;
        if (
          daysSince < coolOffDays &&
          !Config.terminalStatuses.includes(lead.status)
        ) {
          flags.push("cool_off");
        }

        // 🚀 TWEAK: Only flag 3rd Contacts for recycle if they are actively assigned
        if (
          lead.status === "3rd Contact" &&
          daysSince >= coolOffDays &&
          isAssigned
        ) {
          flags.push("needs_recycle");
        }
      }

      // 🚀 THE CLOCK FIX: Check 'lastTouchedOn' first!
      // When a lead is assigned or worked, this updates, resetting the recycle countdown
      // so new agents aren't instantly penalized for an old createdAt date.
      const ref = lead.lastTouchedOn || lead.lastContacted || lead.createdAt;

      if (
        ref &&
        !Config.terminalStatuses.includes(lead.status) &&
        lead.status !== "3rd Contact" &&
        isAssigned // 🚀 TWEAK: Only run the general recycle clock if an agent is holding it
      ) {
        const daysSince = (now - new Date(ref)) / 8640000;
        if (daysSince > recycleAfterDays) flags.push("needs_recycle");
      }

      if (lead.assignedTo && agentCounts[lead.assignedTo] > maxLeadsPerAgent) {
        flags.push("agent_overloaded");
      }

      return Object.assign({}, lead, {
        flags: flags,
        agentLeadCount: agentCounts[lead.assignedTo] || 0,
      });
    });
  }

  // 🛑 Check SharePoint LMSSuspensions list
  async function checkSuspensionStatus(agentEmail) {
    if (!agentEmail) return false;

    try {
      await resolveSiteIds();
      const emailLower = agentEmail.toLowerCase().trim();

      const url =
        base +
        "/sites/" +
        siteIds.leadship +
        "/lists/" +
        Config.sharePoint.lists.suspensionsList +
        "/items?expand=fields";

      const response = await apiFetch(url, { method: "GET" });
      const records = response ? response.value || [] : [];

      if (records.length === 0) return false;

      const now = new Date();

      const parseDuration = (timeString) => {
        if (!timeString) return 0;
        const parts = timeString.split(" ");
        const amount = parseInt(parts[0], 10) || 0;
        const unit = parts[1] ? parts[1].toLowerCase() : "";

        let days = 0;
        if (unit.includes("day")) days = amount;
        if (unit.includes("week")) days = amount * 7;

        return days * 24 * 60 * 60 * 1000;
      };

      // Loop through records
      for (let i = 0; i < records.length; i++) {
        const item = records[i];
        const fields = item.fields;
        const lookupId = fields.AgentLookupId;

        // Skip if the team lead left the Person column blank
        if (!lookupId) continue;

        const createdDate = new Date(item.createdDateTime);
        const durationMs = parseDuration(fields.SuspensionTime);
        const expirationDate = new Date(createdDate.getTime() + durationMs);

        // First, check if this suspension record is actually still active time-wise
        if (now < expirationDate) {
          // 🚀 THE DECODER: It's an active suspension, so we ask SharePoint who this ID belongs to
          const userUrl =
            base +
            "/sites/" +
            siteIds.leadship +
            "/lists/User Information List/items/" +
            lookupId +
            "?expand=fields";

          try {
            const userRes = await apiFetch(userUrl, { method: "GET" });
            if (userRes && userRes.fields) {
              // Extract the email from the hidden system list
              const recordEmail = (userRes.fields.EMail || "")
                .toLowerCase()
                .trim();

              // Now we check if it matches the current logged-in agent!
              if (recordEmail === emailLower) {
                console.warn(
                  `Agent is suspended until ${expirationDate.toLocaleString()}`,
                );
                return expirationDate;
              }
            }
          } catch (userErr) {
            console.warn(
              "Failed to decode AgentLookupId against SharePoint system",
              userErr,
            );
          }
        }
      }

      return false; // Time served!
    } catch (err) {
      console.error("Failed to check suspension status:", err);
      return false;
    }
  }

  function canAgentTakeLead(agentName, leads) {
    const count = leads.filter(function (l) {
      return (
        l.assignedTo === agentName &&
        !Config.terminalStatuses.includes(l.status)
      );
    }).length;
    return count < Config.rules.maxLeadsPerAgent;
  }

  // Recycle a lead — record previous agent, unassign, reset to New
  async function recycleLead(leadId, currentAgent) {
    await resolveSiteIds();

    const lead = State
      ? State.leads.find(function (l) {
          return l.id === leadId;
        })
      : null;

    const prev = lead ? lead.previousAgents || "" : "";
    const newPrev = prev ? prev + ", " + currentAgent : currentAgent;

    // 🚀 THE FIX: Stamp it with today's date instead of null
    const todayDate = new Date().toISOString().split("T")[0];

    // 🛑 THE INTERCEPTOR: Check if this lead has the D2D bounty flag
    const isFlagged =
      lead &&
      (lead.flaggedForExport === true || lead.FlaggedForExport === true);

    // If flagged, lock it away. Otherwise, toss it back in the "New" pool.
    const targetStatus = isFlagged ? "D2D Lead" : "New";

    await updateLead(leadId, {
      Status: targetStatus,
      Agent_x0020_Assigned: null,
      PreviousAgents: newPrev,
      LastTouchedOn: todayDate, // Resets the clock!
    });
  }

  function isInCoolOff(lead) {
    if (lead.status === "New") return false;

    if (!lead.lastContacted) return false;
    const daysSince = (new Date() - new Date(lead.lastContacted)) / 86400000;
    return daysSince < Config.rules.coolOffDays;
  }

  // Count unique leads an agent contacted today.
  // Accepts email directly for reliable matching against activity log.
  function agentContactsToday(agentEmail, activityLog) {
    const emailLower = (agentEmail || "").toLowerCase().trim();
    const uniqueLeads = new Set();

    activityLog.forEach(function (e) {
      // Safety check: Handle both your mapped frontend names AND raw Graph API names
      const entryAgent = (e.agent || e.AgentEmail || "").toLowerCase().trim();
      const actionStr = e.action || e.ActionType || "";
      const leadId = e.leadId || e.LeadID;

      // Is it an actual lead touch?
      const isContact =
        actionStr.startsWith("Status:") ||
        actionStr === "1st Contact" ||
        actionStr === "2nd Contact" ||
        actionStr === "3rd Contact";

      // If it's a contact, by this agent, add the lead ID to the unique Set
      if (isContact && entryAgent === emailLower && leadId) {
        uniqueLeads.add(leadId);
      }
    });

    // A Set automatically prevents duplicates, so its size is the exact number of unique leads worked!
    return uniqueLeads.size;
  }

  // ============================================================
  //  REWARD SHOP (TREMENDOUS API MIDDLEMAN)
  // ============================================================

  async function createRewardOrder(agentEmail, rewardId, cost) {
    if (!agentEmail || !rewardId || !cost) {
      throw new Error("Missing required data to process reward.");
    }

    await resolveSiteIds();

    // Using the leadship site assuming you want to keep this list secure/hidden from the general team site
    const url =
      base +
      "/sites/" +
      siteIds.leadship +
      "/lists/" +
      Config.sharePoint.lists.rewardOrdersList + // 🚀 We will need to add this to Config!
      "/items";

    const payload = {
      fields: {
        Title: rewardId, // e.g., "visa_25"
        AgentEmail: agentEmail,
        Status: "Pending", // Power Automate will watch for "Pending"
      },
    };

    // Use your robust wrapper
    const res = await apiFetch(url, {
      method: "POST",
      body: payload,
    });

    return res;
  }

  return {
    apiFetch,
    getLeads,
    addLead,
    updateLead,
    assignAgent,
    recycleLead,
    getNextLeadForAgent,
    getContractors,
    getActivityLog,
    getActivityLogForToday,
    logActivity,
    getTodaySales,
    getDailyStats,
    resolveSiteIds,
    applyBusinessRules,
    checkSuspensionStatus,
    canAgentTakeLead,
    isInCoolOff,
    agentContactsToday,
    getAgentScores,
    createAgentScore,
    checkLedgerForDuplicate,
    writeLedgerTransaction,
    updateAgentScore,
    createRewardOrder,
    getLeadMessages: SMSApi.getLeadMessages,
    sendTextMessage: SMSApi.sendTextMessage,
    syncSMSHistory: SMSApi.syncSMSHistory,
  };
})();
