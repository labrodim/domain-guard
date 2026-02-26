// DomainGuard - OnMessageSend handler

Office.onReady();

var CFG = (typeof DG_CONFIG !== "undefined") ? DG_CONFIG : {
    internalDomains: [
        "oxfordsla.com",
        "oxfordsla.onmicrosoft.com",
        "metaformalabs.com",
        "metaformalabs.onmicrosoft.com"
    ],
    allowedPartnerDomains: [],
    blockOnError: false
};

function onMessageSendHandler(event) {
    try {
        var item = Office.context.mailbox.item;
        var all = [];
        var pending = 3;

        function checkDomains() {
            var seen = {};
            var unique = [];

            all.forEach(function (r) {
                var domain = (r.emailAddress || "").split("@")[1];
                if (domain) {
                    domain = domain.toLowerCase();
                    if (CFG.internalDomains.indexOf(domain) === -1 &&
                        CFG.allowedPartnerDomains.indexOf(domain) === -1 &&
                        !seen[domain]) {
                        seen[domain] = true;
                        unique.push(domain);
                    }
                }
            });

            if (unique.length > 1) {
                event.completed({
                    allowEvent: false,
                    errorMessage: "Recipients span " + unique.length
                        + " external domains: " + unique.join(", ")
                        + ". Please verify all recipients."
                });
            } else {
                event.completed({ allowEvent: true });
            }
        }

        function done() {
            pending--;
            if (pending === 0) checkDomains();
        }

        ["to", "cc", "bcc"].forEach(function (field) {
            item[field].getAsync(function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
                    all = all.concat(result.value);
                }
                done();
            });
        });
    } catch (e) {
        event.completed({ allowEvent: !CFG.blockOnError });
    }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
