function runLoad() {
    
    objectArray = [];
    rowRule = {};

    Rule1_E = { "SearchString": "E", "SendToLogic": "Substring"};
    Rule1_E.EmailSubject = "URGENTE - Confirmação de Emissão de TED";
    Rule1_E.EmailBody = 
        "Prezados,<br>" +   
        "Favor dar o “de acordo” da emissão  da operação abaixo pelo Tesoureiro ou do Gerente da Unidade, confirmando " + 
        "a sua emissão aonde seus dados estão corretos e também seu valor, e que também o associado possuía saldo para emitir a operação.";

    Rule2_S = { "SearchString": "S", "SendToLogic": "Substring"};
    Rule2_S.EmailSubject = "URGENTE - Confirmação de Emissão de TED";
    Rule2_S.EmailBody = 
        "Prezados,<br>" +   
        "Favor dar o “de acordo” da emissão  da operação abaixo pelo Tesoureiro ou do Gerente da Unidade, confirmando " + 
        "a sua emissão aonde seus dados estão corretos e também seu valor, e que também o associado possuía saldo para emitir a operação.";

    Rule3_I = { "SearchString": "I", "SendToLogic": "Substring"};
    Rule3_I.EmailSubject = "URGENTE - Confirmação de Emissão de TED";
    Rule3_I.EmailBody = 
        "Prezados,<br>" +   
        "Favor dar o “de acordo” da emissão  da operação abaixo pelo Tesoureiro ou do Gerente da Unidade, confirmando " + 
        "a sua emissão aonde seus dados estão corretos e também seu valor, e que também o associado possuía saldo para emitir a operação.";
    
    Rule4_TedConvenios  = { "SearchString": "RIC", "SendToLogic": "Fixed", "SendTo": "convenios_banco@gmail.com.br"};
    Rule4_TedConvenios.EmailSubject = "Confirmação de Emissão de TED - CONVÊNIOS";
    Rule4_TedConvenios.EmailBody = 
        "Prezados,<br>" +
        "Favor dar o “de acordo” da emissão da operação abaixo, confirmando a sua emissão aonde seus dados estão corretos e também seu valor.";
        
    Rule5_Pagfor = { "SearchString": "OP", "SendToLogic": "PagForCNPJ"};
    Rule5_Pagfor.EmailSubject = "Título e-mail: URGENTE- Confirmação de Emissão de TED";
    Rule5_Pagfor.EmailBody = "";

    Rule6_Mafre = { "SearchString": "MAPFRE SEGUROS GERAIS SA", "SendToLogic": "Fixed", "SendTo": "test" };
    Rule6_Mafre.EmailSubject = "Título e-mail: URGENTE- Confirmação de Emissão de TED";
    Rule6_Mafre.EmailBody = "";

    Rule7_Str26 = {}
    Rule7_Str26.EmailSubject = "Título e-mail: URGENTE- Confirmação de Emissão de TED";
    Rule7_Str26.EmailBody = "";

    Rule8_Ag0100 = {}
    Rule8_Ag0100.EmailSubject = "Título e-mail: URGENTE- Confirmação de Emissão de TED";
    Rule8_Ag0100.EmailBody = "";
    
    RuleSet = [Rule1_E, Rule2_S, Rule3_I, Rule4_TedConvenios, Rule5_Pagfor, Rule6_Mafre, Rule7_Str26, Rule8_Ag0100];            
}

function processSpbFile() {

    var spbFileReader = new FileReader();
    spbFileReader.onload = function(event) {

        var jsonEmailsExcel = emailsExcel;
        var jsonEmailsExtra = emailsExtra;

        var spbCont = spbFileReader.result;        

        var object = {};
        var objectEnv = {};
        var objectResp = {};

        ({ object, objectEnv, objectResp } = parseSPB(spbCont, object, objectEnv, objectResp));

        createEmailRequests(jsonEmailsExcel, objectArray);
    }

    var spbFileSelector = document.getElementById('spbFileSelector');
    var spbFile = spbFileSelector.files[0];

    spbFileReader.readAsText(spbFile);
}

function createEmailRequests(jsonEmailsExcel, object) {
    
    console.log("createEmailRequests");
    console.log(object);
    objectArray.forEach(function (objectA) {

        console.log(objectA.ID);
        var email = jsonEmailsExcel.filter(function (el) {
            return el.COOP.toString().startsWith(objectA.ID);
        });

        var subject = 'URGENTE - Confirmação de Emissão de TED - XXXX.XX';
        var body = 'Prezados,\n\nFavor dar o “de acordo” da emissão da operação abaixo pelo Tesoureiro ou do Gerente da Unidade, confirmando a sua emissão aonde seus dados estão corretos e também seu valor, e que também o associado possuía saldo para emitir a operação.';

        body = body + JSON.stringify(objectA, null, 2);

        window.open('mailto:' + email[0].UA + '?subject=' + subject + '&body=' + encodeURIComponent(body));
    });
}

function parseSPB(spbCont, object, objectEnv, objectResp) {

    console.log("parseSPB");
    envHead = "Visualização do XML de Envio";
    respHead = "Visualização do(s) XML(s) de Resposta(s) tipo R1";

    var rowIsEnvio = false;
    var rowIsResposta = false;

    var lines = spbCont.split('\n');
    lines.forEach(function (line) {

        var cleanLine = line.trim();

        var rowIsEnvHead = cleanLine == envHead;
        var rowIsRespHead = cleanLine == respHead;
        var rowIsAttribute = cleanLine.toString().search(":") > -1;

        if (rowIsEnvHead) { 
            // If this is a NEW Envio, 
            // Add OLD envio to the Array and create a new one           
            if (!rowIsEnvio) {
                object["Env"] = objectEnv;
                object["Resp"] = objectResp;
                objectArray.push(object);

                object = {};
                objectEnv = {};
                objectResp = {};
            }

            rowIsEnvio = true;
            rowIsResposta = false;

        } else if (rowIsRespHead) {

            rowIsEnvio = false;
            rowIsResposta = true;

        } else if (rowIsAttribute) {

            var split = cleanLine.split(":");
            var rowAttrName = "" + split[0];
            rowAttrName = rowAttrName.trim();
            var rowAttrValue = "" + split[1];
            rowAttrValue = rowAttrValue.trim();

            var searchStringCollection = {};
            searchStringCollection.Rule1 = Rule1_E.SearchString;
            searchStringCollection.Rule2 = Rule2_S.SearchString;
            searchStringCollection.Rule3 = Rule3_I.SearchString;            
            searchStringCollection.Rule4 = Rule4_TedConvenios.SearchString;
            searchStringCollection.Rule5 = Rule5_Pagfor.SearchString;

            if (rowAttrName == "Número Controle IF") {

                var indexOfLetter = rowAttrValue.search(searchStringCollection.Rule4);
                if (indexOfLetter > -1) {
                    id = "RIC";
                }
                var indexOfLetter = rowAttrValue.search(searchStringCollection.Rule5);
                if (indexOfLetter > -1) {
                    id = "OP";
                }
                if (indexOfLetter <= -1) {
                    indexOfLetter = rowAttrValue.search(searchStringCollection.Rule3);
                    id = rowAttrValue.substring(0, indexOfLetter);
                };
                if (indexOfLetter <= -1) {
                    indexOfLetter = rowAttrValue.search(searchStringCollection.Rule2);
                    id = rowAttrValue.substring(0, indexOfLetter);
                };
                if (indexOfLetter <= -1) {
                    indexOfLetter = rowAttrValue.search(searchStringCollection.Rule1);
                    id = rowAttrValue.substring(0, indexOfLetter);
                };
                if (indexOfLetter <= -1) {
                    console.log("Unable to find separator letter for " + rowAttrValue);
                };

                if (id != "RIC") {
                    if (id.charAt(id.length - 3) == 0) {
                        id = id.substring(id.length - 2, id.length);
                    } else {
                        id = id.substring(id.length - 3, id.length);
                    }
                }

                object.ID = id;
                id = "";

            } else if (rowAttrName == "Agência Debitada" || rowAttrName == "Agência Creditada") {

                if (rowAttrValue.startsWith("0")) {
                    object.ID = rowAttrValue.substr(1, 4) + object.ID;
                } else {
                    object.ID = rowAttrValue + object.ID;
                }

                console.log(object.ID);
            } else if (rowIsEnvio) {
                objectEnv[rowAttrName] = rowAttrValue;
            } else if (rowIsResposta) {
                objectResp[rowAttrName] = rowAttrValue;
            }
        }
    });
    return { object, objectEnv, objectResp };
}
