function Form_onload()
{
var templateId = '790843EA-69F6-DF11-9934-0014C240F49E'; 


function Signature(companyTemplateId) 
{ 
    var sig = this; 
    var emailIframe; 
    var emailBody; 
    var header = GenerateAuthenticationHeader(); 


    sig.TemplateId = companyTemplateId; 


    sig.Load = function() 
    { 
        try 
        { 
                        var xml ='<soap:Envelope ' + 
                        '       xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' + 
                        '       xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' + 
                        '       xmlns:xsd="http://www.w3.org/2001/XMLSchema">' + 
                        header + 
                        '       <soap:Body>' + 
                        '               <Execute xmlns="http://schemas.microsoft.com/crm/2007/ 
WebServices">' + 
                        '                       <Request xsi:type="InstantiateTemplateRequest" 
ReturnDynamicEntities="false">' + 
                        '                       <TemplateId>' +  sig.TemplateId + '</TemplateId>' + 
                        '                       <ObjectType>' + crmForm.ownerid.DataValue[0].typename +'</ 
ObjectType>' + 
                        '                       <ObjectId>' + crmForm.ownerid.DataValue[0].id + '</ObjectId>' + 
                        '                       </Request>' + 
                        '               </Execute>' + 
                        '       </soap:Body>' + 
                        '</soap:Envelope>'; 


//                      alert(xml); 


            var xmlHttpRequest = new ActiveXObject("Msxml2.XMLHTTP"); 
            xmlHttpRequest.open("POST", "/mscrmservices/2007/ 
CrmService.asmx", false); 
            xmlHttpRequest.setRequestHeader("SOAPAction","http:// 
schemas.microsoft.com/crm/2007/WebServices/Execute"); 
            xmlHttpRequest.setRequestHeader("Content-Type", "text/xml; 
charset=utf-8"); 
            xmlHttpRequest.setRequestHeader("Content-Length", 
xml.length); 
            xmlHttpRequest.send(xml); 


            var resultXml = xmlHttpRequest.responseXML; 
            if (xmlHttpRequest.status == 200) 
            { 
                emailBody = resultXml.selectSingleNode("// 
q1:description").text; 
                emailIframeReady(); 
            } 
        } 
        catch(err) 
        { 
            alert(err.description); 
        } 
    } 


    function emailIframeReady() 
    { 
        if (emailIframe.readyState != 'complete') 
        { 
            return; 
        } 


       emailIframe.contentWindow.document.body.innerHTML = emailBody; 
    } 


    emailIframe = document.all.descriptionIFrame; 
    emailIframe.onreadystatechange = emailIframeReady; } 


function OnCrmPageLoad() 
{ 
    if (crmForm.FormType == 1) 
    { 
        var signature = new Signature(templateId); 
        signature.Load(); 
    } 



} 


OnCrmPageLoad();
}
function scheduledend_onchange()
{
var dField = crmForm.all.scheduledend;
if (dField.DataValue != null)
{
var d = new Date(dField.DataValue);
if(d.getHours() == 0)
{
d.setHours(9);
dField.DataValue = d;
}
}
}
