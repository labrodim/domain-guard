Office.onReady(function () {});

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

function getRecipientsAsync(item) {
    return new Promise(function (resolve, reject) {
        var all = [];
        var pending = 3;

        function done() {
            pending--;
            if (pending === 0) resolve(all);
        }

        ["to", "cc", "bcc"].forEach(function (field) {
            item[field].getAsync(function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
                    all = all.concat(result.value);
                }
                done();
            });
        });
    });
}

function onItemSend(event) {
    try {
        getRecipientsAsync(Office.context.mailbox.item).then(function (recipients) {
            var seen = {};
            var unique = [];

            recipients.forEach(function (r) {
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
                Office.context.mailbox.item.notificationMessages.addAsync(
                    "domainGuardWarning",
                    {
                        type: "errorMessage",
                        message: "Warning: Recipients span " + unique.length
                            + " external domains (" + unique.join(", ")
                            + "). Remove this message and re-send to proceed."
                    }
                );
                event.completed({ allowEvent: false });
            } else {
                event.completed({ allowEvent: true });
            }
        });
    } catch (e) {
        event.completed({ allowEvent: !CFG.blockOnError });
    }
}
