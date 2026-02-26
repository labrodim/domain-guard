// ===============================
// Domain Guard Configuration File
// ===============================

const DG_CONFIG = {
    // Base URL where all files are hosted
    baseUrl: "https://metaformalabs.sharepoint.com/sites/oxfordslateam/Shared%20Documents/DomainGuard/",

    // Your internal domains (no warning if only these appear)
    internalDomains: [
        "oxfordsla.com",
        "oxfordsla.onmicrosoft.com",
        "metaformalabs.com",
        "metaformalabs.onmicrosoft.com"
    ],

    // Optional: partner domains you trust (no warning)
    allowedPartnerDomains: [
        // "trustedfirm.com"
    ],

    // Behavior flags
    blockOnError: false,   // if true, blocks send when add-in errors
    requireDominantDomain: false // if true, only warn when an outlier exists
};
