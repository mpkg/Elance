var templateId = '790843EA-69F6-DF11-9934-0014C240F49E'; //'790843EA-69F6-DF11-9934-0014C240F49E';
var emailBody;

function Form_onload() {    
    if (Xrm.Page.ui.getFormType() == 1) {
        var owner = Xrm.Page.getAttribute('ownerid').getValue();

        try {
            SDK.EmailJS.TemplateReqRequest(templateId, owner[0].typename, owner[0].id);
        }
        catch (err) {
            alert(err.description);
        }

        emailIframe = document.all.descriptionIFrame;
        emailIframe.onreadystatechange = emailIframeReady;
    }
}
function scheduledend_onchange() {
    var dField = Xrm.Page.getAttribute("scheduledend");
    if (dField.getValue() != null) {
        var d = new Date(dField.getValue());
        if (d.getHours() == 0) {
            d.setHours(9);
            dField.setValue(d);
        }
    }
}

function emailIframeReady() {
    if (emailIframe.readyState != 'complete') {
        return;
    }
    emailIframe.contentWindow.document.body.innerHTML = emailBody;
}

if (typeof (SDK) == "undefined")
{ SDK = { __namespace: true }; }
//This will establish a more unique namespace for functions in this library. This will reduce the 
// potential for functions to be overwritten due to a duplicate name when the library is loaded.
SDK.EmailJS = {
    _getServerUrl: function () {
        ///<summary>
        /// Returns the URL for the SOAP endpoint using the context information available in the form
        /// or HTML Web resource.
        ///</summary>
        var OrgServicePath = "/XRMServices/2011/Organization.svc/web";
        var serverUrl = "";
        if (typeof GetGlobalContext == "function") {
            var context = GetGlobalContext();
            serverUrl = context.getServerUrl();
        }
        else {
            if (typeof Xrm.Page.context == "object") {
                serverUrl = Xrm.Page.context.getServerUrl();
            }
            else
            { throw new Error("Unable to access the server URL"); }
        }
        if (serverUrl.match(/\/$/)) {
            serverUrl = serverUrl.substring(0, serverUrl.length - 1);
        }
        return serverUrl + OrgServicePath;
    },
    TemplateReqRequest: function (templateID, type, userId) {
        var requestMain = ""
        requestMain += "<s:Envelope xmlns:s='http://schemas.xmlsoap.org/soap/envelope/'>";
        requestMain += "<s:Body><Execute xmlns='http://schemas.microsoft.com/xrm/2011/Contracts/Services' xmlns:i='http://www.w3.org/2001/XMLSchema-instance'>";
        requestMain += "<request i:type='b:InstantiateTemplateRequest' xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts' xmlns:b='http://schemas.microsoft.com/crm/2011/Contracts'>";
        requestMain += "<a:Parameters xmlns:c='http://schemas.datacontract.org/2004/07/System.Collections.Generic'>";
        requestMain += "<a:KeyValuePairOfstringanyType>";
        requestMain += "<c:key>TemplateId</c:key>";
        requestMain += "<c:value i:type='d:guid' xmlns:d='http://schemas.microsoft.com/2003/10/Serialization/'>" + templateID + "</c:value>";
        requestMain += "</a:KeyValuePairOfstringanyType>";
        requestMain += "<a:KeyValuePairOfstringanyType>";
        requestMain += "<c:key>ObjectType</c:key>";
        requestMain += "<c:value i:type='d:string' xmlns:d='http://www.w3.org/2001/XMLSchema'>" + type + "</c:value>";
        requestMain += "</a:KeyValuePairOfstringanyType>";
        requestMain += "<a:KeyValuePairOfstringanyType>";
        requestMain += "<c:key>ObjectId</c:key>";
        requestMain += "<c:value i:type='d:guid' xmlns:d='http://schemas.microsoft.com/2003/10/Serialization/'> " + userId + "</c:value>";
        requestMain += "</a:KeyValuePairOfstringanyType>";
        requestMain += "</a:Parameters>";
        requestMain += "<a:RequestId i:nil='true' />";
        requestMain += "<a:RequestName>InstantiateTemplate</a:RequestName>";
        requestMain += "</request>";
        requestMain += "</Execute>";
        requestMain += "</s:Body>";
        requestMain += "</s:Envelope>";
        var req = new XMLHttpRequest();
        req.open("POST", SDK.EmailJS._getServerUrl(), true)
        // Responses will return XML. It isn't possible to return JSON.
        req.setRequestHeader("Accept", "application/xml, text/xml, */*");
        req.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
        req.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/Execute");
        var successCallback = null;
        var errorCallback = null;
        req.onreadystatechange = function () { SDK.EmailJS.TemplateReqResponse(req); };
        req.send(requestMain);
    },
    TemplateReqResponse: function (req) {
        ///<summary>
        /// Recieves the assign response
        ///</summary>
        ///<param name="req" Type="XMLHttpRequest">
        /// The XMLHttpRequest response
        ///</param>
        ///<param name="successCallback" Type="Function">
        /// The function to perform when an successfult response is returned.
        /// For this message no data is returned so a success callback is not really necessary.
        ///</param>
        ///<param name="errorCallback" Type="Function">
        /// The function to perform when an error is returned.
        /// This function accepts a JScript error returned by the _getError function
        ///</param>
        if (req.readyState == 4) {
            if (req.status == 200) {
                var resultXml = req.responseXML;
                x = resultXml.getElementsByTagName("a:KeyValuePairOfstringanyType");
                for (i = 0; i < x.length; i++) {
                    if (x[i].childNodes[0].textContent == 'description') {
                        emailBody = x[i].childNodes[1].textContent;
                    }

                }
                emailIframeReady();
            }
            else {
                alert(SDK.EmailJS._getError(req.responseXML));
            }
        }
    },
    _getError: function (faultXml) {
        ///<summary>
        /// Parses the WCF fault returned in the event of an error.
        ///</summary>
        ///<param name="faultXml" Type="XML">
        /// The responseXML property of the XMLHttpRequest response.
        ///</param>
        var errorMessage = "Unknown Error (Unable to parse the fault)";
        if (typeof faultXml == "object") {
            try {
                var bodyNode = faultXml.firstChild.firstChild;
                //Retrieve the fault node
                for (var i = 0; i < bodyNode.childNodes.length; i++) {
                    var node = bodyNode.childNodes[i];
                    //NOTE: This comparison does not handle the case where the XML namespace changes
                    if ("s:Fault" == node.nodeName) {
                        for (var j = 0; j < node.childNodes.length; j++) {
                            var faultStringNode = node.childNodes[j];
                            if ("faultstring" == faultStringNode.nodeName) {
                                errorMessage = faultStringNode.textContent;
                                break;
                            }
                        }
                        break;
                    }
                }
            }
            catch (e) { };
        }
        return new Error(errorMessage);
    },
    __namespace: true
};
