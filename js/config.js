// Raimak LMS - App Configuration v3.0

const Config = {
  azure: {
    clientId: "ad1b153f-8b6a-4f1c-8ab2-58fcf03cf5c2",
    tenantId: "39e14190-0b23-4ecd-99f9-606ad1215881",
    redirectUri: window.location.origin + window.location.pathname,
  },

  sharePoint: {
    hostname: "raimak.sharepoint.com",
    sites: {
      leadship: "sites/RaimakLeadship",
      team: "TeamSite",
    },
    lists: {
      activityLog: "E3F03088-D099-47E7-B548-0FC4D727E93C",
      contractorList: "bd5df38a-9cb6-411d-87e8-3e79934213d3",
      leadsList: "D698470F-673E-474F-8EE7-D02D54814E31",
      agentScores: "242d56aa-dfcd-4399-919b-0ce608ec932c",
      agentScoresLedger: "7c6f1604-bd68-4ceb-9f0c-9271faf5e0b8",
      suspensionsList: "d30048fb-8f64-4ed9-9719-6fec2aaca9a9",
      rewardOrdersList: "06A1DC4C-F120-4BF6-9E53-F5ECE25A58BD",
      smsHistory: "a5add1b8-c36b-4f41-953b-df06099d82a8",
      vzlmsConfig: "08bd72de-270c-4df3-ac3c-f1a6be612b3b",
    },
    graphBase: "https://graph.microsoft.com/v1.0",
  },

  rules: {
    coolOffDays: 2,
    maxLeadsPerAgent: 99999,
    maxContactsPerDay: 99999,
    recycleAfterDays: 30,
    appVersion: "1.0",
  },

  // Pipeline statuses
  leadStatuses: [
    "New",
    "1st Contact",
    "2nd Contact",
    "3rd Contact",
    "Do Not Call",
    "Already has VZ",
    "Yes",
  ],

  // Terminal statuses — removed from agent queue, admin only
  terminalStatuses: [
    "Do Not Call",
    "Sold",
    "FNQ",
    "Already has VZ",
    "TDM",
    "D2D Lead",
  ],

  // TDM is kicked back to admin (D2D only)
  adminOnlyStatuses: ["TDM"],

  soldStatus: "Sold",

  // Lead types
  leadTypes: ["OFS", "MLR", "Forced"],

  // Current products options — alphabetical
  currentProducts: [
    "Home Phone + Internet + VAS",
    "Internet",
    "Internet + Phone",
    "Internet + TV",
    "Internet + TV + Phone",
    "Internet + VAS",
    "Other",
    "Phone",
    "TV",
    "TV + Phone",
  ],

  leadSources: [
    "Web Form",
    "Referral",
    "Cold Call",
    "Email Campaign",
    "Social Media",
    "Trade Show",
    "Other",
  ],

  roles: {
    admins: [
      "RichardZacker@raimak.com",
      "B.Hinesley@raimak.com",
      "S.Balleste@raimak.com",
      "N.Caldwell@raimak.com",
      "C.Scarrett@raimak.com",
      "m.mcalpine@raimak.com",
      "J.Scroggins@raimak.com",
      "antoinette.bickel@raimak.com",
      "a.cooper@raimak.com",
    ],
  },

  scopes: ["Sites.ReadWrite.All", "User.Read"],

  salesFeedInterval: 45000,
};
