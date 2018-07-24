// Notification Template Form Scripts

function Form_OnLoad() {

    var hideDueDate = false;
    var dueDate = Xrm.Page.getAttribute("didd_duedate").getValue();
    var duration = Xrm.Page.getAttribute("didd_duration").getValue();
    var dueDateModifier = Xrm.Page.getAttribute("didd_duedatemodifier").getValue();
    var dueDateModifierAttribute = Xrm.Page.getAttribute("didd_duedatemodifierattribure").getValue();

    if (dueDate != null && dueDate != undefined) {
        Xrm.Page.getControl('didd_duedate').setVisible(true);

        Xrm.Page.getAttribute("didd_duration").setValue(null);
        Xrm.Page.getAttribute("didd_duedatemodifier").setValue(null);
        Xrm.Page.getAttribute("didd_duedatemodifierattribure").setValue(null);

        Xrm.Page.getControl('didd_duration').setVisible(false);
        Xrm.Page.getControl('didd_duedatemodifier').setVisible(false);
        Xrm.Page.getControl('didd_duedatemodifierattribure').setVisible(false);
    }
    else if (duration != null || dueDateModifier != null || dueDateModifierAttribute != null) {
        Xrm.Page.getControl('didd_duedate').setVisible(false);
        Xrm.Page.getControl('didd_duration').setVisible(true);
        Xrm.Page.getControl('didd_duedatemodifier').setVisible(true);
        Xrm.Page.getControl('didd_duedatemodifierattribure').setVisible(true);

        Xrm.Page.getAttribute("didd_duedate").setValue(null);
    }
    else if (dueDate == null || duration == null || dueDateModifier == null || dueDateModifierAttribute == null) {
        Xrm.Page.getControl('didd_duedate').setVisible(true);
        Xrm.Page.getControl('didd_duration').setVisible(true);
        Xrm.Page.getControl('didd_duedatemodifier').setVisible(true);
        Xrm.Page.getControl('didd_duedatemodifierattribure').setVisible(true);
    }


}
function LoadAttributes(selectedObj, selectControl) {
    var entityName = selectedObj.value;
    var query = "?$select=LogicalName&$expand=Attributes($select=LogicalName,DisplayName)&$filter=LogicalName eq '" + entityName + "'";
    var resultEntity = TnXrmUtilities.WebAPI.RetrieveMultipleRecordsSync("EntityDefinitions", query, function errorCallBackFun() { });
    resultEntity[0].Attributes.sort(sortOn("LogicalName"));
    var displayName = [];
    var logicalName = [];
    displayName.length = resultEntity[0].Attributes.length;
    logicalName.length = resultEntity[0].Attributes.length;
    for (var i = 0, j = 0; i < resultEntity[0].Attributes.length ; i++, j++) {
        if (resultEntity[0].Attributes[i] != null) {
            if (resultEntity[0].Attributes[i].DisplayName.UserLocalizedLabel != null) {
                if (resultEntity[0].Attributes[i].DisplayName.UserLocalizedLabel.Label != null) {
                    displayName[j] = resultEntity[0].Attributes[i].DisplayName.UserLocalizedLabel.Label;
                    logicalName[j] = resultEntity[0].Attributes[i].LogicalName;;

                    option = document.createElement("option");
                    option.setAttribute("value", logicalName[j]);
                    option.innerHTML = displayName[j];
                    //gt dom element from xrm
                    var destination = window.parent.Xrm.Page.getControl("WebResource_notificationtemplatemetadata").getObject().contentWindow.document;
                    destination.getElementById(selectControl).appendChild(option);

                }
            }
        }
    }
}
function LoadMetaData() {
    var query = "?$select=LogicalName,DisplayName,IsCustomizable&$filter=IsCustomizable/Value eq true and IsActivity eq false&$count=true";
    var resultEntity = TnXrmUtilities.WebAPI.RetrieveMultipleRecordsSync("EntityDefinitions", query, errorCallBackFun);
    resultEntity.sort(sortOn("LogicalName"));
    var displayName = [];
    var logicalName = [];
    displayName.length = resultEntity.length;
    logicalName.length = resultEntity.length;
    for (var i = 0; i < resultEntity.length ; i++) {
        if (resultEntity[i] != null) {
            if (resultEntity[i].DisplayName.LocalizedLabels.length != 0) {
                if (resultEntity[i].DisplayName.LocalizedLabels[0] != null) {
                    displayName[i] = resultEntity[i].DisplayName.LocalizedLabels[0].Label;
                    logicalName[i] = resultEntity[i].LogicalName;

                    option = document.createElement("option");
                    option.setAttribute("value", logicalName[i]);
                    option.innerHTML = displayName[i];
                    //gt dom element from xrm
                    var destination = window.parent.Xrm.Page.getControl("WebResource_notificationtemplatemetadata").getObject().contentWindow.document;
                    destination.getElementById("SelectEntity").appendChild(option);
                }
            }
        }
    }

    function errorCallBackFun(param) {
        alert("Operation Failed !!" + param);
    }

    function successCallBackFun(param) {
        alert("Operation Success !!");
    }
}
function sortOn(property) {
    return function (a, b) {
        if (a[property] < b[property]) {
            return -1;
        } else if (a[property] > b[property]) {
            return 1;
        } else {
            return 0;
        }
    }
}
function SetAttribute(SelectEntity, SelectAttribute, SelectSetAttribute) {
    var entityAttribField = SelectSetAttribute.value;
    if (SelectAttribute.value != null && SelectAttribute.value != 0) {
        var entityAttribValue = "{" + SelectEntity.value + "!" + SelectAttribute.value + "}";
    } else {
        var entityAttribValue = SelectEntity.value;
    }
    var existingValue = window.parent.Xrm.Page.getAttribute(entityAttribField).getValue() != null ? window.parent.Xrm.Page.getAttribute(entityAttribField).getValue() : "";
    window.parent.Xrm.Page.getAttribute(entityAttribField).setValue(existingValue + " " + entityAttribValue);
}