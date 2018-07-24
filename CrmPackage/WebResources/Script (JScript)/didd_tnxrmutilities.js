/*jslint browser: true, devel: true, plusplus: true, newcap: true  */
/*global Xrm */
/*jslint nomen: true */
/*global GetGlobalContext*/
/*global $*/
/*global CrmEncodeDecode*/
/*global Oncomplete*/
/*jslint maxlen: 250 */
var TnXrmUtilities,
    notificationsArea;

var Months30 = [4, 6, 9, 11];
var Months31 = [1, 3, 5, 7, 8, 10, 12];

if (TnXrmUtilities === undefined || TnXrmUtilities === null) {
    /// <summary>
    /// TnXrmUtilities.Common for common functions.
    /// TnXrmUtilities.REST for all REST endpoints functions.
    /// </summary>
    TnXrmUtilities = {};

    TnXrmUtilities.Common = (function () {
        "use strict";
        /// <summary>
        /// Display an alert message using the XRM alert dialog function
        /// </summary>
        /// <param name="message" type="string">
        /// The message to be displayed to the user
        /// </param>
        /// <returns type="void" />
        var alertMessage = function (messagecode) {
            var message = getMessage(messagecode);
            if (Xrm.Utility !== undefined &&
                Xrm.Utility !== null &&
                Xrm.Utility.alertDialog !== undefined &&
                Xrm.Utility.alertDialog !== null) {
                Xrm.Utility.alertDialog(message);
            } else {
                alert(message);
            }
        },

            /// <summary>
            /// Sets Lookup Value.
            /// </summary>
            /// <param name="fieldName" type="string">
            /// Lookup Field schema name.
            /// </param>
            /// <param name="id" type="string">
            /// Lookup field id.
            /// </param>
            /// <param name="name" type="string">
            /// Name of the lokkup field.
            /// </param>
            /// <param name="entityType" type="string">
            /// type of the field.
            /// </param>
            /// <returns type="void" />
            setLookupValue = function (fieldName, id, name, entityType) {
                if (fieldName != null) {
                    if (id.indexOf('{') == -1)
                        id = '{' + id;
                    if (id.indexOf('}') == -1)
                        id = id + '}';
                    id = id.toUpperCase();

                    var lookupValue = new Array();
                    lookupValue[0] = new Object();
                    lookupValue[0].id = id;
                    lookupValue[0].name = name;
                    lookupValue[0].entityType = entityType;
                    Xrm.Page.getAttribute(fieldName).setValue(lookupValue);
                }
            },

            /// <summary>
            /// Get Text of Lookup Value.
            /// </summary>
            /// <param name="lookupSchemaName" type="string">
            /// Lookup Field schema name.
            /// </param>
            /// <returns type="string" />
            getLookupTextValue = function (lookupSchemaName) {
                var lookupObj = Xrm.Page.getAttribute(lookupSchemaName);
                var lookupRecordName;
                if (lookupObj != null && lookupObj != undefined) {
                    var lookupObjValue = lookupObj.getValue();
                    if (lookupObjValue != null) {
                        var lookupEntityType = lookupObjValue[0].entityType;
                        var lookupRecordGuid = lookupObjValue[0].id;
                        lookupRecordName = lookupObjValue[0].name;
                    }
                }
                return lookupRecordName;
            },

            /// <summary>
            /// Compares two look-up text values.
            /// </summary>
            /// <param name="lookup1SchemaName" type="string">
            /// Lookup Field schema name of the First Field
            /// </param>
            /// <param name="lookup2SchemaName" type="string">
            /// Lookup Field schema name to be compared with
            /// </param>
            /// <returns type="boolean" />
            areLookupValuesSame = function (lookup1SchemaName, lookup2SchemaName) {
                var lookup1Obj = Xrm.Page.getAttribute(lookup1SchemaName);
                var lookup2Obj = Xrm.Page.getAttribute(lookup2SchemaName);
                var lookup1RecordName, lookup2RecordName;
                if (lookup1Obj != null && lookup1Obj != undefined) {
                    var lookup1ObjValue = lookup1Obj.getValue();
                    if (lookup1ObjValue != null) {
                        var lookup1EntityType = lookup1ObjValue[0].entityType;
                        var lookup1RecordGuid = lookup1ObjValue[0].id;
                        lookup1RecordName = lookup1ObjValue[0].name;
                    }
                }
                if (lookup2Obj != null && lookup2Obj != undefined) {
                    var lookup2ObjValue = lookup2Obj.getValue();
                    if (lookup2ObjValue != null) {
                        var lookup2EntityType = lookup2ObjValue[0].entityType;
                        var lookup2RecordGuid = lookup2ObjValue[0].id;
                        lookup2RecordName = lookup2ObjValue[0].name;
                    }
                }
                if (lookup1RecordName != null && lookup2RecordName != null && lookup1RecordName == lookup2RecordName)
                    return true;
                else return false;
            },

            /// <summary>
            /// Gives the current Date time.
            /// </summary>
            /// <returns type= "returns the Current date time" />
            currentDateTime = function () {
                var currentDateTime = new Date();
                var date = currentDateTime.getDate();
                var month = currentDateTime.getUTCMonth() + 1;
                var year = currentDateTime.getFullYear();
                var hours = currentDateTime.getHours();
                var minutes = currentDateTime.getMinutes();
                var ampm = " AM";

                if (hours > 11) {
                    ampm = " PM";
                    hours = hours - 12;
                }

                if (hours == 0) {
                    hours = 12;
                }

                if (minutes < 10) {
                    var time = hours.toString() + ":0" + minutes.toString() + ampm;
                } else {
                    var time = hours.toString() + ":" + minutes.toString() + ampm;
                }

                var now = month + "/" + date + "/" + year + " " + time;
                return now;
            },

            /// <summary>
            /// Format the Phone, fax  and validation on phone & fax.
            /// </summary>
            /// <param name="phone" type="string">
            /// Phone Field Value.
            /// </param>
            /// <param name="phoneSchema" type="string">
            /// Phone schema name.
            /// </param>
            /// <param name="uniqueId" type="string">
            /// uniqueId of the Form.
            /// </param>
            /// <returns type="void" />
            phoneValidFormat = function (phone, phoneSchema, uniqueId) {
                var clearPhone1 = null;
                var regExp = null;
                var formtedPhone = null;
                var indexOfReturn = phoneSchema.indexOf("fax");

                if (phone != null) {
                    clearPhone1 = phone.replace("-", "");
                    clearPhone1 = clearPhone1.replace("-", "");
                    clearPhone1 = clearPhone1.replace("(", "");
                    clearPhone1 = clearPhone1.replace(")", "");
                    clearPhone1 = clearPhone1.replace("X", "");
                    clearPhone1 = clearPhone1.replace(" ", "");
                    clearPhone1 = clearPhone1.replace(" ", "");
                    clearPhone1 = clearPhone1.replace(" ", "");

                    phone = clearPhone1;
                    if ((phone.length >= 11) && (phone.length <= 17)) {
                        regExp = new RegExp("^[0-9]{11,17}$");
                    }
                    else {
                        regExp = new RegExp("^[0-9]{10}$");
                    }                     
                    if (regExp.test(phone) == false && indexOfReturn != -1) {
                        // Fax Number entered is invalid.
                        setControlNotification("407", phoneSchema, uniqueId);
                    } else if (regExp.test(phone) == false && indexOfReturn == -1) {
                        // Phone Number entered is invalid.
                        setControlNotification("006", phoneSchema, uniqueId);
                    }

                    if (phone.length == 10 && regExp.test(phone) == true) {
                        clearControlNotification(phoneSchema, uniqueId);
                        formtedPhone = "(" + phone.substr(0, 3) + ") " + phone.substr(3, 3) + "-" + phone.substr(6, 4);
                        Xrm.Page.getAttribute(phoneSchema).setValue(formtedPhone);
                    }

                    if ((phone.length >= 11 && phone.length <= 17) && regExp.test(phone) == true) {
                        clearControlNotification(phoneSchema, uniqueId);
                        formtedPhone = "(" + phone.substr(0, 3) + ") " + phone.substr(3, 3) + "-" + phone.substr(6, 4) + " " + "X" + " " + phone.substr(10, (phone.length - 10));
                        Xrm.Page.getAttribute(phoneSchema).setValue(formtedPhone);
                    }


                } else {
                    clearControlNotification(phoneSchema, uniqueId);
                }
            },

            /// <summary>
            /// Format the zip code and validation on zip code.
            /// </summary>
            /// <param name="ZipSchema" type="string">
            /// Zipcode schema name.
            /// </param>
            /// <returns type="void" />
            zipCodeValidFormat = function (ZipSchema, uniqueId) {
                var zip = Xrm.Page.getAttribute(ZipSchema).getValue();
                var clearZip = null;

                if (zip != null) {
                    clearZip = zip.replace("-", "");
                    clearZip = clearZip.replace("-", "");
                    zip = clearZip;

                    var regExp = new RegExp("^[0-9]{9}$");
                    var regExp2 = new RegExp("^[0-9]{5}$");
                    if (regExp.test(zip) == false && regExp2.test(zip) == false) {
                        // Zip Code format must be either xxxxx or xxxxx-xxxx.
                        setControlNotification("229", ZipSchema, uniqueId);
                    }

                    if (regExp.test(zip) == true) {
                        clearControlNotification(ZipSchema, uniqueId);
                        var formtedZip = zip.substr(0, 5) + "-" + zip.substr(5, 4);
                        Xrm.Page.getAttribute(ZipSchema).setValue(formtedZip);
                    }

                    if (regExp2.test(zip) == true) {
                        clearControlNotification(ZipSchema, uniqueId);
                    }
                } else {
                    clearControlNotification(ZipSchema, uniqueId);
                }
            },

            /// <summary
            //Furture Days should not be allowed.
            /// </summary>
            /// </param>
            /// <returns type="void" />
            futureDatesNotAllowed = function (schemaName, uniqueId, errorCode) {
                var currentDateTime = new Date();
                var field = Xrm.Page.getAttribute(schemaName);
                var fieldValue = null;

                if (field != null) {
                    fieldValue = field.getValue();
                }

                if (fieldValue != null && fieldValue > currentDateTime) {
                    //Date cannot be a future date.
                    setControlNotification(errorCode, schemaName, uniqueId);
                } else {
                    clearControlNotification(schemaName, uniqueId);
                }
            },

            /// <summary
            //It returns True if person status is "Active"
            /// </summary>
            /// <param name="personSchema" type="string">
            /// Person schema name.
            /// </param>
            /// <returns type="void" />
            checkForPersonStatus = function (personSchema) {

                var personName = null;
                var personValue = null;
                var personId = null;
                var flag = 0;
                var person = Xrm.Page.data.entity.attributes.get(personSchema);

                //Check State of Person
                if (person != null && person != undefined && person.getValue() != null && person.getValue().length > 0) {
                    personValue = person.getValue();
                    if (personValue != null && personValue != undefined) {
                        personName = person.getValue()[0].name;
                        personId = person.getValue()[0].id;
                    }
                }

                if (person == null || person == undefined || person.getValue() == null) {
                    return true;
                }

                if (personId != null) {
                    var odataSelect1 = "statuscode";
                    var results = TnXrmUtilities.WebAPI.RetrieveRecordSync(personId.replace("{", "").replace("}", ""), "leads", odataSelect1);

                    if (results != null && results != undefined) {
                        var status = results.statuscode;

                        if (status != null && status != 273310000) {
                            //Person record status is not Active.
                            flag = 1;
                        }
                    }

                    if (flag == 0) {
                        return true;
                    } else {
                        return false;
                    }
                }
            },

            checkPersonStatus = function (personSchema, statusValues) {

                var personName = null;
                var personValue = null;
                var personId = null;
                var flag = 0;
                var person = Xrm.Page.data.entity.attributes.get(personSchema);

                //Check State of Person
                if (person != null && person != undefined && person.getValue() != null && person.getValue().length > 0) {
                    personValue = person.getValue();
                    if (personValue != null && personValue != undefined) {
                        personName = person.getValue()[0].name;
                        personId = person.getValue()[0].id;
                    }
                }

                if (person == null || person == undefined || person.getValue() == null) {
                    return true;
                }

                if (personId != null) {
                    var odataSelect1 = "statuscode";
                    var results = TnXrmUtilities.WebAPI.RetrieveRecordSync(personId.replace("{", "").replace("}", ""), "leads", odataSelect1);

                    if (results != null && results != undefined) {
                        var status = results.statuscode;

                        if (status != null && status != 273310000 && status != 273310005 && status != 273310007 && status != 273310010) {
                            //Person record status is not Active or Waitlisted or Pending Enrollment.
                            flag = 1;
                        }

                        //$.each(statusValues, function (index, value) {
                        //    if (status != statusValues[index]) {
                        //        flag = 1;
                        //    }
                        //});

                    }

                    if (flag == 0) {
                        return true;
                    } else {
                        return false;
                    }
                }
            },

            /// <summary
            // Email Format and validation.
            /// </summary>
            /// <param name="emailSchema" type="string">
            /// Email schema name.
            /// </param>
            /// <param name="uniqueId" type="string">
            /// Unique Id of the form.
            /// </param>
            /// <returns type="void" />
            emailValidation = function (emailSchema, uniqueId) {
                var email = Xrm.Page.getAttribute(emailSchema);

                if (email != null && email != undefined) {
                    var emailValue = email.getValue();

                    if (emailValue != null) {
                        var atpos = emailValue.indexOf("@");
                        var dotpos = emailValue.lastIndexOf(".");
                        if (atpos < 1 || dotpos < atpos + 2 || dotpos + 2 >= emailValue.length) {
                            // Email Address entered is invalid.
                            setControlNotification("408", emailSchema, uniqueId);
                        } else {
                            clearControlNotification(emailSchema, uniqueId);
                        }
                    }
                }
            },

            /// <summary
            /// IQ Test Validation
            /// </summary>
            /// <param name="IqSchema" type="string">
            /// IQ schema name.
            /// </param>
            /// <param name="uniqueId" type="string">
            /// Unique Id of the form.
            /// </param>
            /// <returns type="void" />
            iqValidation = function (iqSchema, uniquieId) {
                var iqTest = Xrm.Page.getAttribute(iqSchema).getValue();
                var regExp = null;

                //if IQ validation field is not null
                if (iqTest != null) {
                    if (iqTest.length == 1) {
                        regExp = new RegExp("^[0-9]{1}$");
                    } else if (iqTest.length == 2) {
                        regExp = new RegExp("^[0-9]{2}$");
                    } else {
                        regExp = new RegExp("^[0-9]{3}$");
                    }

                    if (regExp.test(iqTest) == false) {
                        //IQ Score must be numeric and cannot have more than three digits
                        setControlNotification("242", iqSchema, uniquieId);
                    } else {
                        clearControlNotification(iqSchema, uniquieId);
                    }
                }
            },


            /// <summary>
            /// Display a confirmation message to the user.
            /// </summary>
            /// <param name="message" type="string">
            /// The message to be displayed to the user
            /// </param>
            /// <param name="yesCallback" type="method">
            /// Method to be invoked when the user selects "Yes"
            /// </param>
            /// <param name="noCallback" type="method">
            /// Method to be invoked when the user selects "No"
            /// </param>
            /// <returns type="void" />
            confirmMessage = function (message, yesCallback, noCallback) {
                if (Xrm.Utility !== undefined &&
                    Xrm.Utility !== null &&
                    Xrm.Utility.confirmDialog !== undefined &&
                    Xrm.Utility.confirmDialog !== null) {
                    Xrm.Utility.confirmDialog(message, yesCallback, noCallback);
                } else {
                    if (confirm(message)) {
                        yesCallback();
                    } else {
                        noCallback();
                    }
                }
            },

            //Set CrmForm Notification.
            setFormNotification = function (messagecode, level, uniqueId) {
                var message = getMessage(messagecode);
                if (level === 1) {
                    //Error
                    Xrm.Page.ui.setFormNotification(message, "ERROR", uniqueId);
                }
                if (level === 2) {
                    //Info
                    Xrm.Page.ui.setFormNotification(message, "INFO", uniqueId);
                }
                if (level === 3) {
                    //Warning
                    Xrm.Page.ui.setFormNotification(message, "WARNING", uniqueId);
                }
            },

            //Clear CrmForm Notification.
            clearFormNotification = function (uniqueId) {
                Xrm.Page.ui.clearFormNotification(uniqueId);
            },

            //Set CRMControl Notification
            setControlNotification = function (messagecode, attributeName, uniqueId) {
                var message = getMessage(messagecode);
                Xrm.Page.getControl(attributeName).setNotification(message, uniqueId)
            },

            //Clear CRMControl Notification
            clearControlNotification = function (attributeName, uniqueId) {
                Xrm.Page.getControl(attributeName).clearNotification(uniqueId);
            },


            /// <summary>
            /// Check if two guids are equal
            /// </summary>
            /// <param name="guid1" type="string">
            /// A string represents a guid
            /// </param>
            /// <param name="guid2" type="string">
            /// A string represents a guid
            /// </param>
            /// <returns type="boolean" />
            guidsAreEqual = function (guid1, guid2) {
                var isEqual;
                if (guid1 === null ||
                    guid2 === null ||
                    guid1 === undefined ||
                    guid2 === undefined) {
                    isEqual = false;
                } else {
                    isEqual = guid1.replace(/[{}]/g, "").toLowerCase() === guid2.replace(/[{}]/g, "").toLowerCase();
                }

                return isEqual;
            },

            /// <summary>
            /// Enable a field by the name
            /// </summary>
            /// <param name="fieldName" type="string">
            /// The name of the field to be enabled
            /// </param>
            /// <returns type="void" />
            enableField = function (fieldName) {
                if (Xrm.Page.getControl(fieldName) !== null) {
                    Xrm.Page.getControl(fieldName).setDisabled(false);
                }
            },

            /// <summary>
            /// Disable a field by the name
            /// </summary>
            /// <param name="fieldName" type="string">
            /// The name of the field to be disabled
            /// </param>
            /// <returns type="void" />
            disableField = function (fieldName) {
                if (Xrm.Page.getControl(fieldName) !== null) {
                    Xrm.Page.getControl(fieldName).setDisabled(true);
                }
            },

            /// <summary>
            /// Show a field by the name
            /// </summary>
            /// <param name="fieldName" type="string">
            /// The name of the field to be shown
            /// </param>
            /// <returns type="void" />
            showField = function (fieldName) {
                if (Xrm.Page.getControl(fieldName) !== null) {
                    Xrm.Page.getControl(fieldName).setVisible(true);
                }
            },

            /// <summary>
            /// Hide a field by the name
            /// </summary>
            /// <param name="fieldName" type="string">
            /// The name of the field to be hidden
            /// </param>
            /// <returns type="void" />
            hideField = function (fieldName) {
                if (Xrm.Page.getControl(fieldName) !== null) {
                    Xrm.Page.getControl(fieldName).setVisible(false);
                }
            },

            /// <summary>
            /// Updates the requirement level of a field
            /// </summary>
            /// <param name="fieldName" type="string">
            /// Name of the field
            /// </param>
            /// <param name="levelName" type="string">
            /// Name of the requirement level. [none, recommended, required] (Case Sensitive)
            /// </param>
            /// <returns type="void" />
            updateRequirementLevel = function (fieldName, levelName) {
                if (Xrm.Page.getControl(fieldName) !== null) {
                    Xrm.Page.getAttribute(fieldName).setRequiredLevel(levelName);
                }
            },

            /// <summary>
            /// Alert the error message if occurred
            /// </summary>
            /// <param name="error" type="error">
            /// Object of the JavaScript error
            /// </param>
            /// <returns type="void" />
            showError = function (error) {
                alertMessage(error.message);
            },

            /// <summary>
            /// Checking the role exist for the user
            /// </summary>
            /// <param name="roleName" type="role">
            /// object of collection of roles
            /// </param>
            /// <returns type="bool" />
            userHasRole = function (roleName) {
                /// <summary>Checks whether user belongs to the role passed or not.</summary>
                /// <param name="roleName" type="String">Pass rolename of type String.</param>
                /// <returns type="Boolean">Returns true or false.</returns>
                var currentUserRoles = Xrm.Page.context.getUserRoles(),
                    appendRoles, userRoles,
                    roles = roleName.split(','),
                    query,
                    innerQueryTemplate = "roleid eq {GUID}",
                    innerRoleQuery = "name eq '{RoleName}'",
                    innerQuery = "",
                    i, j, userRoleId, response = null,
                    appendRoles;
                appendRoles = "";
                for (j = 0; j < roles.length; j++) {
                    userRoles = roles[j];
                    if (appendRoles == "") {
                        appendRoles = innerRoleQuery.replace("{RoleName}", userRoles);
                    } else {
                        appendRoles += " or " + innerRoleQuery.replace("{RoleName}", userRoles);
                    }
                }


                //innerRoleQuery = innerRoleQuery.replace("{innerRoleQuery}", appendRoles);

                for (i = 0; i < currentUserRoles.length; i++) {
                    userRoleId = currentUserRoles[i];
                    if (innerQuery === "") {
                        innerQuery = innerQueryTemplate.replace("{GUID}", userRoleId);
                    } else {
                        innerQuery += " or " + innerQueryTemplate.replace("{GUID}", userRoleId);
                    }
                }

                query = "$filter=(" + appendRoles + ") and (" + innerQuery + ")";
                response = TnXrmUtilities.WebAPI.RetrieveMultipleRecordsSync("roles", query);

                if (response !== undefined && response !== null && response.length > 0) {
                    return true;
                }

                return false;
            },

            /// <summary>
            /// Calculate the days between two dates
            /// </summary>
            /// <param name="datetime1" type="DateTime">
            /// The first / early date to be calculated
            /// </param>
            /// <param name="datetime2" type="DateTime">
            /// The second / later date to e calculated
            /// </param>
            /// <returns type="int" />
            calculateDaysBetween = function (datetime1, datetime2) {
                // The number of milliseconds in one day
                var oneDay = 1000 * 60 * 60 * 24,
                    // Convert both dates to milliseconds
                    date1Ms = datetime1.getTime(),
                    date2Ms = datetime2.getTime(),
                    // Calculate the difference in milliseconds
                    differenceMs = Math.abs(date1Ms - date2Ms);

                // Convert back to days and return
                return Math.round(differenceMs / oneDay);
            },

            /// <summary>
            /// Disable all controls in a tab by tab number.
            /// </summary>
            /// <param name="tabControlNo" type="int">
            /// The number of the tab
            /// </param>
            /// <returns type="void" />
            disableAllControlsInTab = function (tabControlNo) {
                var tabControl = Xrm.Page.ui.tabs.get(tabControlNo);
                if (tabControl !== undefined && tabControl !== null) {
                    Xrm.Page.ui.controls.forEach(
                        function (control) {
                            if (control.getParent() !== null && control.getParent().getParent() !== null && control.getParent().getParent() === tabControl && control.getControlType() !== "subgrid") {
                                control.setDisabled(true);
                            }
                        }
                    );
                }
            },

            ///common function to calculate age in months for infant and also form the display format exactly
            calculateAge = function (dateOfBirth, flagForMonths) {

                var today, age, monthDiff, i, year, yearInLoop, result;
                monthDiff = 0;
                today = new Date();
                year = today.getFullYear();
                yearInLoop = dateOfBirth.getFullYear();

                age = calculateDiffDaysOnly(today, dateOfBirth);

                for (i = (dateOfBirth.getMonth() + 1) ; i <= 12; i++) {
                    if (i > (today.getMonth() + 1) && yearInLoop === year) {
                        break;
                    }


                    if (Months30.indexOf(i) > -1 && age >= 30) {
                        monthDiff = monthDiff + 1;
                        age = age - 30;
                    }

                    if (Months31.indexOf(i) > -1 && age >= 31) {
                        monthDiff = monthDiff + 1;
                        age = age - 31;
                    } else if (Months31.indexOf(i) > -1 && age < 31) {
                        break;
                    }

                    if (i === 2 && yearInLoop % 4 === 0 && age >= 29) {
                        monthDiff = monthDiff + 1;
                        age = age - 29;
                    } else if (i === 2 && age >= 28) {
                        monthDiff = monthDiff + 1;
                        age = age - 28;
                    }


                    if (i === 12) {
                        yearInLoop = yearInLoop + 1;
                        i = 0;
                    }

                }
                result = (flagForMonths) ? monthDiff : monthDiff + "m" + " " + age + "d";


                return result;
            },

            // common function to set Age Format for Child , Infant and Others
            //accepted values for ageGroup are "woman" and all other category names
            setAgeFormat = function (dateOfBirth, ageGroup) {

                var ageInFormat, years, months, today;
                today = new Date();
                if (ageGroup !== null && ageGroup !== undefined) {
                    ageGroup = ageGroup.toUpperCase();
                }
                if (ageGroup === Constants.child.toUpperCase()) {
                    months = Math.floor(calculateAge(dateOfBirth, true));
                    if (months <= 24) {
                        ageInFormat = calculateAge(dateOfBirth, false);
                    } else {
                        years = Math.floor(calculateAge(dateOfBirth, true) / 12);
                        ageInFormat = years + "y" + " " + months + "m";
                    }
                } else if (ageGroup === Constants.infant.toUpperCase()) {
                    ageInFormat = calculateAge(dateOfBirth, false);
                } else {
                    years = Math.floor(calculateAge(dateOfBirth, true) / 12);
                    ageInFormat = years + "y";
                }
                return ageInFormat;
            },

            /// <summary>
            /// Disable all controls in a section by section label.
            /// </summary>
            /// <param name="sectionLabel" type="string">
            /// The label of the section
            /// </param>
            /// <returns type="void" />
            disableAllControlsInSection = function (sectionLabel) {
                var tabs = Xrm.Page.ui.tabs,
                    i,
                    tab,
                    sections,
                    j,
                    section,
                    sectionlength,
                    tablength;
                for (i = 0, tablength = tabs.getLength() ; i < tablength; i++) {
                    tab = tabs.get(i);
                    sections = tab.sections;
                    for (j = 0, sectionlength = sections.getLength() ; j < sectionlength; j++) {
                        section = sections.get(j);
                        if (section.getLabel().toLowerCase() === sectionLabel.toLowerCase()) {
                            Xrm.Page.ui.controls.forEach(
                                function (control) {
                                    if (control.getParent() !== null && control.getParent().getLabel() === sectionLabel && control.getControlType() !== "subgrid") {
                                        control.setDisabled(true);
                                    }
                                }
                            );
                            break;
                        }
                    }
                }
            };


        /// <summary>
        /// Validating Email address.
        /// <param name="field" type="string">
        /// Will check email address format and reture true false accordingly
        /// </param>
        /// </summary>
        /// <returns type="boolean" />
        function validateEmail(field) {
            var regex = /^[a-zA-Z0-9._\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,5}$/;
            return (regex.test(field)) ? true : false;
        }

        /// <summary>
        /// Validating Email address.
        /// <param name="emailcntl" type="string">
        /// Will check the multiple email address and reture true false accordingly
        /// </param>
        /// </summary>
        /// <returns type="boolean" />
        function validateMultipleEmailsCommaSeparated(emailcntl) {
            var result, value = emailcntl,
                i;
            if (value.length > 0) {
                result = value.split(";");

                for (i = 0; i < result.length; i++) {
                    if (result[i].length > 0) {
                        if (!validateEmail(result[i])) {
                            return false;
                        }
                    } else {
                        return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// convert current date to EDM format to use it in Odata Query.
        /// </param>
        /// </summary>
        /// <returns type="datetime" />

        function formatDateEDM() {
            var d = new Date(),
                month = d.getMonth() + 1,
                day = d.getDate(),
                year = d.getFullYear();

            if (month.toString().length < 2) {
                month = '0' + month;
            }
            if (day.toString().length < 2) {
                day = '0' + day;
            }

            return [year, month, day].join('-') + 'T12:00:00';
        }

        /// <summary>
        /// disable the form fields
        ///<param>
        /// pass bool value to disable the fields 
        /// </param>
        /// </summary>
        var doesControlHaveAttribute = function (control) {
            var controlType = control.getControlType();
            return controlType != "iframe" && controlType != "webresource" && controlType != "subgrid";
        },
            disableFormFields = function (onOff) {
                Xrm.Page.ui.controls.forEach(function (control, index) {
                    if (doesControlHaveAttribute(control)) {
                        control.setDisabled(onOff);
                    }
                });
            },
            onChangeHomePhoneNumber = function (PhoneFieldName) {
                /// <summary>Triggers on change of phone number field.</summary>
                /// <param name="PhoneFieldName" type="sigle line of text">Pass field name</param>
                //Validating Home Phone Number
                var phoneNumber, regex;
                phoneNumber = Xrm.Page.data.entity.attributes.get(PhoneFieldName).getValue();
                if (null !== phoneNumber) {
                    regex = /^[0-9]{3}-[0-9]{3}-[0-9]{4}$/;
                    if ((!(isNaN(phoneNumber)) && phoneNumber.length === 10) || regex.test(phoneNumber)) {
                        Xrm.Page.getControl(PhoneFieldName).clearNotification(1);
                    } else {
                        //Xrm.Page.data.entity.attributes.get(PhoneFieldName).setValue(null);
                        Xrm.Page.getControl(PhoneFieldName).setFocus();
                        Xrm.Page.getControl(PhoneFieldName).setNotification("Invalid Number :- Formats Allowed are XXX-XXX-XXXX or XXXXXXXXXX", 1);

                    }
                } else {
                    Xrm.Page.getControl(PhoneFieldName).clearNotification(1);
                }
            },
            showHideSections = function (tabName, commaSeperatedSectionNames, state) {
                if (Xrm.Page.ui.tabs.get(tabName) != null) {
                    var tabSections = Xrm.Page.ui.tabs.get(tabName).sections.get();
                    var givenSectionNames = commaSeperatedSectionNames.split(',');

                    for (var i = 0; i < tabSections.length; i++) {
                        for (var j = 0; j < givenSectionNames.length; j++) {

                            if (tabSections[i].getName() == givenSectionNames[j]) {
                                tabSections[i].setVisible(state);
                            }
                        }
                    }
                }
            },
            showHideTab = function (tabName, hide) {
                if (Xrm.Page.ui.tabs.get(tabName) != null) {
                    Xrm.Page.ui.tabs.get(tabName).setVisible(!hide);
                }
            },

            getTeamId = function (teamname) {
                var returnValue = null;
                var odataSelect = "$filter=name eq '" + teamname + "'";
                var requestResults = TnXrmUtilities.WebAPI.RetrieveMultipleRecordsSync("teams", odataSelect);

                if (requestResults != undefined && requestResults !== null && requestResults.length > 0) {
                    returnValue = requestResults[0].teamid;
                }

                return returnValue;
            },

            userHasTeam = function (teamname) {
                var teamnames, requestResults, returnValue = false,
                    teamnames = teamnames.split(','),
                    currentUserId = Xrm.Page.context.getUserId(),
                    j,
                    odataSelect = "$filter=systemuserid eq " + currentUserId.replace('{', '').replace('}', '');
                requestResults = TnXrmUtilities.WebAPI.RetrieveMultipleRecordsSync("teammemberships", odataSelect);
                for (j = 0; j <= teamnames.length; j++) {
                    if (requestResults !== undefined && requestResults !== null && requestResults.length > 0) {
                        var teamId = getTeamId(teamname[j]);
                        for (var i = 0; i < requestResults.length; i++) {
                            if (requestResults[i].teamid == teamId) {
                                returnValue = true;
                            }
                        }
                    } else {
                        returnValue = false;
                    }
                }

                return returnValue;
            },

            saveAndComplete = function (statuscode, statecode, updateIEACompletedOn) {
                //You are about to complete this record. Once this action takes place,the record becomes locked and will no longer be available for edit. Continue?(059) 
                var ok = confirm(getMessage("059"));

                if (ok != true) {
                    return;
                }

                // Get the status
                var recordStatus = Xrm.Page.getAttribute("statuscode");

                // if status is not null
                if (recordStatus != null) {
                    // set the status as "Complete"
                    recordStatus.setValue(statuscode);
                }

                // Get the record state
                var recordState = Xrm.Page.getAttribute("statecode");

                // if statecode is not null
                if (recordState != null) {
                    // set the record state to inactive.
                    recordState.setValue(statecode);
                }

                if (updateIEACompletedOn != null || updateIEACompletedOn != undefined) {
                    if (updateIEACompletedOn == true) {
                        var completedOnDate = Xrm.Page.getAttribute("didd_completedon");
                        if (completedOnDate != null && completedOnDate.getValue() == null) {
                            // Set Waitlist Eligibility Completion Date as Current Date.
                            completedOnDate.setValue(new Date());
                        }
                    }
                }
                // Save the record
                Xrm.Page.data.entity.save("saveandclose");
            },

            getExceptionMessage = function (messagecode, event) {
                try {
                    var message = getMessage(messagecode);
                    if (message != null && message != undefined) {
                        return message;
                    }
                    else {
                        message = getErrorMessage(event);
                        return message;
                    }
                }
                catch (error) {
                    message = getErrorMessage(event);
                    return message;
                }
            },

            getErrorMessage = function (event) {
                var errorMsg;
                if (event == "OnSave") {
                    errorMsg = "An error occurred while saving the record."
                }
                else if (event == "OnChange") {
                    errorMsg = "An error occurred on field value change."
                }
                else if (event == "OnLoad") {
                    errorMsg = "An error occurred while loading the record."
                }
                else {
                    errorMsg = "An error occurred."
                }

                return errorMsg;
            },

            getMessage = function (messagecode) {
                var returnValue = null;
                var odataSelect = "$filter=didd_messagecode eq '" + messagecode + "'";
                var requestResults = TnXrmUtilities.WebAPI.RetrieveMultipleRecordsSync("didd_messages", odataSelect);

                if (requestResults != undefined && requestResults !== null && requestResults.length > 0) {
                    returnValue = requestResults[0].didd_messagedescription;
                }

                return returnValue;
            };
        // Toolkit's public static members
        return {
            AlertMessage: alertMessage,
            ConfirmMessage: confirmMessage,
            EnableField: enableField,
            DisableField: disableField,
            ShowField: showField,
            HideField: hideField,
            UpdateRequiredLevel: updateRequirementLevel,
            CalculateDaysBetween: calculateDaysBetween,
            ShowError: showError,
            GuidsAreEqual: guidsAreEqual,
            DisableAllControlsInTab: disableAllControlsInTab,
            DisableAllControlsInSection: disableAllControlsInSection,
            CalculateAge: calculateAge,
            SetAgeFormat: setAgeFormat,
            ValidateMultipleEmailsCommaSeparated: validateMultipleEmailsCommaSeparated,
            FormatDateEDM: formatDateEDM,
            DisableFormFields: disableFormFields,
            OnChangeHomePhoneNumber: onChangeHomePhoneNumber,
            GetExceptionMessage: getExceptionMessage,
            ShowHideSections: showHideSections,
            ShowHideTab: showHideTab,
            UserHasRole: userHasRole,
            GetTeamId: getTeamId,
            UserHasTeam: userHasTeam,
            GetMessage: getMessage,
            SetFormNotification: setFormNotification,
            ClearFormNotification: clearFormNotification,
            SetControlNotification: setControlNotification,
            ClearControlNotification: clearControlNotification,
            SetLookupValue: setLookupValue,
            GetLookupTextValue: getLookupTextValue,
            AreLookupValuesSame: areLookupValuesSame,
            CurrentDateTime: currentDateTime,
            PhoneValidFormat: phoneValidFormat,
            ZipCodeValidFormat: zipCodeValidFormat,
            FutureDatesNotAllowed: futureDatesNotAllowed,
            CheckForPersonStatus: checkForPersonStatus,
            EmailValidation: emailValidation,
            IqValidation: iqValidation,
            SaveAndComplete: saveAndComplete,
            CheckPersonStatus: checkPersonStatus,
        };
    }());

    TnXrmUtilities.SOAP = (function () {
        "use strict";
        //Private Properties
        //Start

        ///<summary>
        /// Private function to the context object.
        ///</summary>
        ///<returns>Context</returns>
        var context = function () {
            ///<summary>
            /// Private function to the context object.
            ///</summary>
            ///<returns>Context</returns>
            var oContext = null;
            if (typeof window.GetGlobalContext != "undefined") {
                oContext = window.GetGlobalContext();
            } else {
                if (typeof Xrm != "undefined") {
                    if (typeof Xrm != "undefined") {
                        oContext = Xrm.Page.context;
                    } else {
                        if (typeof window.parent.Xrm != "undefined") {
                            oContext = window.parent.Xrm.Page.context;
                        }
                    }
                } else {
                    throw new Error("Context is not available.");
                }
            }
            return oContext;
        },

            retrieveEntityMetadata = function (entityFilters, logicalName, retrieveIfPublished, callback) {
                ///<summary>
                /// Sends an synchronous/asynchronous RetreiveEntityMetadata Request to retrieve a particular entity metadata in the system
                ///</summary>
                ///<returns>Entity Metadata</returns>
                ///<param name="entityFilters" type="String">
                /// The filter string available to filter which data is retrieved. Case Sensitive filters [Entity,Attributes,Privileges,Relationships]
                /// Include only those elements of the entity you want to retrieve in the array. Retrieving all parts of all entities may take significant time.
                ///</param>
                ///<param name="logicalName" type="String">
                /// The string of the entity logical name
                ///</param>
                ///<param name="retrieveIfPublished" type="Boolean">
                /// Sets whether to retrieve the metadata that has not been published.
                ///</param>
                ///<param name="callBack" type="Function">
                /// The function that will be passed through and be called by a successful response.
                /// This function also used as an indicator if the function is synchronous/asynchronous
                ///</param>

                entityFilters = isArray(entityFilters) ? entityFilters : [entityFilters];
                var entityFiltersString = "";
                for (var iii = 0, templength = entityFilters.length; iii < templength; iii++) {
                    entityFiltersString += encodeValue(entityFilters[iii]) + " ";
                }

                var request = [
                    "<request i:type=\"a:RetrieveEntityRequest\" xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\">",
                    "<a:Parameters xmlns:b=\"http://schemas.datacontract.org/2004/07/System.Collections.Generic\">",
                    "<a:KeyValuePairOfstringanyType>",
                    "<b:key>EntityFilters</b:key>",
                    "<b:value i:type=\"c:EntityFilters\" xmlns:c=\"http://schemas.microsoft.com/xrm/2011/Metadata\">", encodeValue(entityFiltersString), "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                    "<a:KeyValuePairOfstringanyType>",
                    "<b:key>MetadataId</b:key>",
                    "<b:value i:type=\"c:guid\"  xmlns:c=\"http://schemas.microsoft.com/2003/10/Serialization/\">", encodeValue("00000000-0000-0000-0000-000000000000"), "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                    "<a:KeyValuePairOfstringanyType>",
                    "<b:key>RetrieveAsIfPublished</b:key>",
                    "<b:value i:type=\"c:boolean\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">", encodeValue(retrieveIfPublished.toString()), "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                    "<a:KeyValuePairOfstringanyType>",
                    "<b:key>LogicalName</b:key>",
                    "<b:value i:type=\"c:string\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">", encodeValue(logicalName), "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                    "</a:Parameters>",
                    "<a:RequestId i:nil=\"true\" />",
                    "<a:RequestName>RetrieveEntity</a:RequestName>",
                    "</request>"
                ].join("");

                var async = !!callback;

                return doRequesting(request, "Execute", async, function (resultXml) {
                    var response = selectNodes(resultXml, "//b:value");

                    var results = [];
                    for (var i = 0, ilength = response.length; i < ilength; i++) {
                        var a = objectifyNode(response[i]);
                        a._type = "EntityMetadata";
                        results.push(a);
                    }

                    if (!async)
                        return results;
                    else
                        callback(results);
                    // ReSharper disable NotAllPathsReturnValue
                });
                // ReSharper restore NotAllPathsReturnValue
            },

            doRequesting = function (soapBody, requestType, async, internalCallback) {
                async = async || false;

                // Wrap the Soap Body in a soap:Envelope.
                var soapXml = ["<soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>",
                    "<soap:Body>",
                    "<", requestType, " xmlns='http://schemas.microsoft.com/xrm/2011/Contracts/Services' xmlns:i='http://www.w3.org/2001/XMLSchema-instance'>", soapBody, "</", requestType, ">",
                    "</soap:Body>",
                    "</soap:Envelope>"
                ].join("");

                var req = new XMLHttpRequest();
                req.open("POST", orgServicePath(), async);
                req.setRequestHeader("Accept", "application/xml, text/xml, */*");
                req.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
                req.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/" + requestType);

                //IE10
                try {
                    req.responseType = 'msxml-document';
                } catch (e) { }

                if (async) {
                    req.onreadystatechange = function () {
                        if (req.readyState == 4 /* complete */) {
                            req.onreadystatechange = null; //Addresses potential memory leak issue with IE
                            if (req.status === 200) { // "OK"          
                                var doc = req.responseXML;
                                try {
                                    setSelectionNamespaces(doc);
                                } catch (e) { }
                                internalCallback(doc);
                            } else {
                                getError(req);
                            }
                        }
                    };

                    req.send(soapXml);
                } else {
                    req.send(soapXml);
                    if (req.status == 200) {
                        var doc = req.responseXML;
                        try {
                            setSelectionNamespaces(doc);
                        } catch (e) { }
                        var result = doc;
                        return !!internalCallback ? internalCallback(result) : result;
                    } else {
                        getError(req);
                    }
                }
                // ReSharper disable NotAllPathsReturnValue
            },

            isArray = function (input) {
                return input.constructor.toString().indexOf("Array") != -1;
            },

            encodeValue = function (value) {
                // Handle GUIDs wrapped in braces
                if (typeof value == typeof "" && value.slice(0, 1) == "{" && value.slice(-1) == "}") {
                    value = value.slice(1, -1);
                }

                // ReSharper disable QualifiedExpressionMaybeNull
                return (typeof value === "object" && value.getTime)
                    // ReSharper restore QualifiedExpressionMaybeNull
                    ?
                    encodeDate(value) :
                    crmXmlEncode(value);
            },

            encodeDate = function (dateTime) {
                return dateTime.getFullYear() + "-" +
                    padNumber(dateTime.getMonth() + 1) + "-" +
                    padNumber(dateTime.getDate()) + "T" +
                    padNumber(dateTime.getHours()) + ":" +
                    padNumber(dateTime.getMinutes()) + ":" +
                    padNumber(dateTime.getSeconds());
            },

            crmXmlEncode = function (s) {
                // ReSharper restore UnusedLocals
                // ReSharper disable UsageOfPossiblyUnassignedValue
                // ReSharper disable ExpressionIsAlwaysConst
                if ('undefined' === typeof s || 'unknown' === typeof s || null === s) return s;
                    // ReSharper restore ExpressionIsAlwaysConst
                    // ReSharper restore UsageOfPossiblyUnassignedValue
                else if (typeof s != "string") s = s.toString();
                return innerSurrogateAmpersandWorkaround(s);
            },

            innerSurrogateAmpersandWorkaround = function (s) {
                var buffer = '';
                var c0;
                var cnt;
                var cntlength;
                for (cnt = 0, cntlength = s.length; cnt < cntlength; cnt++) {
                    c0 = s.charCodeAt(cnt);
                    if (c0 >= 55296 && c0 <= 57343)
                        if (cnt + 1 < s.length) {
                            var c1 = s.charCodeAt(cnt + 1);
                            if (c1 >= 56320 && c1 <= 57343) {
                                buffer += "CRMEntityReferenceOpen" + ((c0 - 55296) * 1024 + (c1 & 1023) + 65536).toString(16) + "CRMEntityReferenceClose";
                                cnt++;
                            } else
                                buffer += String.fromCharCode(c0);
                        } else buffer += String.fromCharCode(c0);
                    else buffer += String.fromCharCode(c0);
                }
                s = buffer;
                buffer = "";
                for (cnt = 0, cntlength = s.length; cnt < cntlength; cnt++) {
                    c0 = s.charCodeAt(cnt);
                    if (c0 >= 55296 && c0 <= 57343)
                        buffer += String.fromCharCode(65533);
                    else buffer += String.fromCharCode(c0);
                }
                s = buffer;
                s = htmlEncode(s);
                s = s.replace(/CRMEntityReferenceOpen/g, "&#x");
                s = s.replace(/CRMEntityReferenceClose/g, ";");
                return s;
            },

            htmlEncode = function (s) {
                if (s === null || s === "" || s === undefined) return s;
                for (var count = 0, buffer = "", hEncode = "", cnt = 0, sLength = s.length; cnt < sLength; cnt++) {
                    var c = s.charCodeAt(cnt);
                    if (c > 96 && c < 123 || c > 64 && c < 91 || c === 32 || c > 47 && c < 58 || c === 46 || c === 44 || c === 45 || c === 95)
                        buffer += String.fromCharCode(c);
                    else buffer += "&#" + c + ";";
                    if (++count === 500) {
                        hEncode += buffer;
                        buffer = "";
                        count = 0;
                    }
                }
                if (buffer.length) hEncode += buffer;
                return hEncode;
            },

            selectNodes = function (node, xPathExpression) {
                if (typeof (node.selectNodes) != "undefined") {
                    return node.selectNodes(xPathExpression);
                } else {
                    var output = [];
                    var xPathResults = node.evaluate(xPathExpression, node, nsResolver, XPathResult.ANY_TYPE, null);
                    var result = xPathResults.iterateNext();
                    while (result) {
                        output.push(result);
                        result = xPathResults.iterateNext();
                    }
                    return output;
                }
            },

            selectSingleNode = function (node, xpathExpr) {
                if (typeof (node.selectSingleNode) != "undefined") {
                    return node.selectSingleNode(xpathExpr);
                } else {
                    var xpe = new XPathEvaluator();
                    var results = xpe.evaluate(xpathExpr, node, nsResolver, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                    return results.singleNodeValue;

                }
            },

            selectSingleNodeText = function (node, xpathExpr) {
                var x = selectSingleNode(node, xpathExpr);
                if (isNodeNull(x)) {
                    return null;
                }
                if (typeof (x.text) != "undefined") {
                    return x.text;
                } else {
                    return x.textContent;
                }
            },

            setSelectionNamespaces = function (doc) {
                var namespaces = [
                    "xmlns:s='http://schemas.xmlsoap.org/soap/envelope/'",
                    "xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts'",
                    "xmlns:i='http://www.w3.org/2001/XMLSchema-instance'",
                    "xmlns:b='http://schemas.datacontract.org/2004/07/System.Collections.Generic'",
                    "xmlns:c='http://schemas.microsoft.com/xrm/2011/Metadata'",
                    "xmlns:ser='http://schemas.microsoft.com/xrm/2011/Contracts/Services'"
                ];
                doc.setProperty("SelectionNamespaces", namespaces.join(" "));
            },

            arrayElements = ["Attributes",
                "ManyToManyRelationships",
                "ManyToOneRelationships",
                "OneToManyRelationships",
                "Privileges",
                "LocalizedLabels",
                "Options",
                "Targets"
            ],

            isMetadataArray = function (elementName) {
                for (var i = 0, ilength = arrayElements.length; i < ilength; i++) {
                    if (elementName === arrayElements[i]) {
                        return true;
                    }
                }
                return false;
            },

            getNodeName = function (node) {
                if (typeof (node.baseName) != "undefined") {
                    return node.baseName;
                } else {
                    return node.localName;
                }
            },

            objectifyNode = function (node) {
                //Check for null
                if (node.attributes != null && node.attributes.length == 1) {
                    if (node.attributes.getNamedItem("i:nil") != null && node.attributes.getNamedItem("i:nil").nodeValue == "true") {
                        return null;
                    }
                }

                //Check if it is a value
                if ((node.firstChild != null) && (node.firstChild.nodeType == 3)) {
                    var nodeName = getNodeName(node);

                    switch (nodeName) {
                        //Integer Values        
                        case "ActivityTypeMask":
                        case "ObjectTypeCode":
                        case "ColumnNumber":
                        case "DefaultFormValue":
                        case "MaxValue":
                        case "MinValue":
                        case "MaxLength":
                        case "Order":
                        case "Precision":
                        case "PrecisionSource":
                        case "LanguageCode":
                            return parseInt(node.firstChild.nodeValue, 10);
                            // Boolean values
                        case "AutoRouteToOwnerQueue":
                        case "CanBeChanged":
                        case "CanTriggerWorkflow":
                        case "IsActivity":
                        case "IsActivityParty":
                        case "IsAvailableOffline":
                        case "IsChildEntity":
                        case "IsCustomEntity":
                        case "IsCustomOptionSet":
                        case "IsDocumentManagementEnabled":
                        case "IsEnabledForCharts":
                        case "IsGlobal":
                        case "IsImportable":
                        case "IsIntersect":
                        case "IsManaged":
                        case "IsReadingPaneEnabled":
                        case "IsValidForAdvancedFind":
                        case "CanBeSecuredForCreate":
                        case "CanBeSecuredForRead":
                        case "CanBeSecuredForUpdate":
                        case "IsCustomAttribute":
                        case "IsPrimaryId":
                        case "IsPrimaryName":
                        case "IsSecured":
                        case "IsValidForCreate":
                        case "IsValidForRead":
                        case "IsValidForUpdate":
                        case "IsCustomRelationship":
                        case "CanBeBasic":
                        case "CanBeDeep":
                        case "CanBeGlobal":
                        case "CanBeLocal":
                            return (node.firstChild.nodeValue === "true") ? true : false;
                            //OptionMetadata.Value and BooleanManagedProperty.Value and AttributeRequiredLevelManagedProperty.Value
                        case "Value":
                            //BooleanManagedProperty.Value
                            if ((node.firstChild.nodeValue === "true") || (node.firstChild.nodeValue == "false")) {
                                return (node.firstChild.nodeValue == "true") ? true : false;
                            }
                            //AttributeRequiredLevelManagedProperty.Value
                            if (
                                (node.firstChild.nodeValue == "ApplicationRequired") ||
                                (node.firstChild.nodeValue == "None") ||
                                (node.firstChild.nodeValue == "Recommended") ||
                                (node.firstChild.nodeValue == "SystemRequired")
                            ) {
                                return node.firstChild.nodeValue;
                            } else {
                                //OptionMetadata.Value
                                return parseInt(node.firstChild.nodeValue, 10);
                            }
                            // ReSharper disable JsUnreachableCode
                            break;
                            // ReSharper restore JsUnreachableCode   
                            //String values        
                        default:
                            return node.firstChild.nodeValue;
                    }

                }

                //Check if it is a known array
                if (isMetadataArray(getNodeName(node))) {
                    var arrayValue = [];
                    for (var iii = 0, tempLength = node.childNodes.length; iii < tempLength; iii++) {
                        var objectTypeName;
                        if ((node.childNodes[iii].attributes != null) && (node.childNodes[iii].attributes.getNamedItem("i:type") != null)) {
                            objectTypeName = node.childNodes[iii].attributes.getNamedItem("i:type").nodeValue.split(":")[1];
                        } else {

                            objectTypeName = getNodeName(node.childNodes[iii]);
                        }

                        var b = objectifyNode(node.childNodes[iii]);
                        b._type = objectTypeName;
                        arrayValue.push(b);

                    }

                    return arrayValue;
                }

                //Null entity description labels are returned as <label/> - not using i:nil = true;
                if (node.childNodes.length == 0) {
                    return null;
                }

                //Otherwise return an object
                var c = {};
                if (node.attributes.getNamedItem("i:type") != null) {
                    c._type = node.attributes.getNamedItem("i:type").nodeValue.split(":")[1];
                }
                for (var i = 0, ilength = node.childNodes.length; i < ilength; i++) {
                    if (node.childNodes[i].nodeType == 3) {
                        c[getNodeName(node.childNodes[i])] = node.childNodes[i].nodeValue;
                    } else {
                        c[getNodeName(node.childNodes[i])] = objectifyNode(node.childNodes[i]);
                    }

                }
                return c;
            },

            xmlParser = function (txt) {
                ///<summary>
                /// cross browser responseXml to return a XML object
                ///</summary>
                var xmlDoc = null;
                try {
                    // code for Mozilla, Firefox, Opera, etc.
                    if (window.DOMParser) {
                        // ReSharper disable InconsistentNaming
                        var parser = new DOMParser();
                        // ReSharper restore InconsistentNaming
                        xmlDoc = parser.parseFromString(txt, "text/xml");
                    } else // Internet Explorer
                    {
                        xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
                        xmlDoc.async = false;
                        xmlDoc.loadXML(txt);
                    }
                } catch (e) {
                    alert("Cannot convert the XML string to a cross-browser XML object.");
                }

                return xmlDoc;
            },

            xmlToString = function (responseXml) {
                var xmlString = '';
                try {
                    if (responseXml != null) {
                        //IE
                        if (window.ActiveXObject) {
                            xmlString = responseXml.xml;
                        }
                            // code for Mozilla, Firefox, Opera, etc.
                        else {
                            // ReSharper disable InconsistentNaming
                            xmlString = (new XMLSerializer()).serializeToString(responseXml);
                            // ReSharper restore InconsistentNaming
                        }
                    }
                } catch (e) {
                    alert("Cannot convert the XML to a string.");
                }
                return xmlString;
            },

            xrmValue = function (sType, sValue) {
                this.type = sType;
                this.value = sValue;
            },

            fetch = function (fetchXml, callback) {
                alert("This method is obsolete and will be removed soon");
                ///<summary>
                /// Sends synchronous/asynchronous request to do a fetch request.
                ///</summary>
                ///<param name="fetchXml" type="String">
                /// A JavaScript String with properties corresponding to the fetchXml
                /// that are valid for fetch operations.
                /// </param>
                ///<param name="callback" type="Function">
                /// A Function used for asynchronous request. If not defined, it sends a synchronous request.
                /// </param>
                var msgBody = "<query i:type='a:FetchExpression' xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts'>" +
                    "<a:Query>" +
                    ((typeof window.CrmEncodeDecode != 'undefined') ? window.CrmEncodeDecode.CrmXmlEncode(fetchXml) : crmXmlEncode(fetchXml)) +
                    "</a:Query>" +
                    "</query>";
                var async = !!callback;

                return doRequest(msgBody, "RetrieveMultiple", !!callback, function (resultXml) {
                    var fetchResult = $(resultXml).find("a\\:Entities").eq(0);

                    var results = [];

                    for (var i = 0; i < fetchResult.children().length; i++) {
                        var entity = new businessEntity();
                        var deserializeEntity = fetchResult.children().eq(i);
                        entity.deserialize(deserializeEntity);
                        results[i] = entity;
                    }

                    if (!async)
                        return results;
                    else
                        callback(results);
                    // ReSharper disable NotAllPathsReturnValue
                });
                // ReSharper restore NotAllPathsReturnValue
            },
            doRequest = function (soapBody, requestType, async, internalCallback) {
                async = async || false;

                // Wrap the Soap Body in a soap:Envelope.
                var soapXml = ["<soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>",
                    "<soap:Body>",
                    "<", requestType, " xmlns='http://schemas.microsoft.com/xrm/2011/Contracts/Services' xmlns:i='http://www.w3.org/2001/XMLSchema-instance'>", soapBody, "</", requestType, ">",
                    "</soap:Body>",
                    "</soap:Envelope>"
                ].join("");

                var req = getXhr();
                req.open("POST", orgServicePath(), async);
                req.setRequestHeader("Accept", "application/xml, text/xml, */*");
                req.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
                req.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/" + requestType);

                req.send(soapXml);

                if (async) {
                    req.onreadystatechange = function () {
                        if (req.readyState == 4) { // "complete"
                            if (req.status == 200) { // "OK"
                                internalCallback(processResponse(req.responseXML, req.responseText));
                            } else {
                                throw new Error("HTTP-Requests ERROR: " + req.statusText);
                            }
                        }
                    };
                } else {
                    var result = processResponse(req.responseXML);
                    return !!internalCallback ? internalCallback(result) : result;
                }
                // ReSharper disable NotAllPathsReturnValue
            },
            orgServicePath = function () {
                ///<summary>
                /// Private function to return the path to the organization service.
                ///</summary>
                ///<returns>String</returns>
                return getServerUrl() + "/XRMServices/2011/Organization.svc/web";
            },
            getServerUrl = function () {
                ///<summary>
                /// Private function to return the server URL from the context
                ///</summary>
                ///<returns>String</returns>
                var serverUrl = context().getClientUrl();
                if (serverUrl.match(/\/$/)) {
                    serverUrl = serverUrl.substring(0, serverUrl.length - 1);
                }
                return serverUrl;
            },
            processResponse = function (responseXml, responseText) {
                if (typeof jQuery == 'undefined') {
                    throw new Error('jQuery is not loaded.\nPlease ensure that jQuery is included\n as webresource in the form load.');
                }

                if (responseXml === null || responseXml.xml === null || responseXml.xml === "") {
                    if (responseText !== null && responseText !== "")
                        throw new Error(responseText);
                    else
                        throw new Error("No response received from the server. ");
                }

                // Report the error if occurred
                var error = $(responseXml).children("error").text();
                var faultString = $(responseXml).children("faultstring").text();

                if (error != '' || faultString != '') {
                    throw new Error(error !== null ? $(responseXml).children('description').text() : faultString);
                }

                // Load responseXML and return as an XML object
                var xmlDoc = xmlParser(xmlToString(responseXml));
                return xmlDoc;
            },
            getXhr = function () {
                ///<summary>
                /// Get an instance of XMLHttpRequest for all browers
                ///</summary>
                if (XMLHttpRequest) {
                    // Chrome, Firefox, IE7+, Opera, Safari
                    // ReSharper disable InconsistentNaming
                    return new XMLHttpRequest();
                    // ReSharper restore InconsistentNaming
                }
                // IE6
                try {
                    // The latest stable version. It has the best security, performance,
                    // reliability, and W3C conformance. Ships with Vista, and available
                    // with other OS's via downloads and updates.
                    return new ActiveXObject('MSXML2.XMLHTTP.6.0');
                } catch (e) {
                    try {
                        // The fallback.
                        return new ActiveXObject('MSXML2.XMLHTTP.3.0');
                    } catch (e) {
                        alert('This browser is not AJAX enabled.');
                        return null;
                    }
                }
            },

            ///<summary>
            /// Private function to return the server URL from the context
            ///</summary>
            ///<returns>String</returns>
            getClientUrl = function () {
                var clientUrl = context().getClientUrl();
                return clientUrl;
            },
            executeRequest = function (_XML, Message) {
                alert("This method is obsolete and will be removed soon");
                var _ResultXML = null,
                    msg,
                    errorCount,
                    xmlhttp = new XMLHttpRequest();
                xmlhttp.open("POST", getClientUrl() + "/XRMServices/2011/Organization.svc/web", false);
                xmlhttp.setRequestHeader("Accept", "application/xml, text/xml, */*");
                xmlhttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
                xmlhttp.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/Execute");
                xmlhttp.send(_XML);
                _ResultXML = xmlhttp.responseXML;
                errorCount = $(_ResultXML).find("a\\:error").length;
                if (errorCount !== 0) {
                    msg = $(_ResultXML).find("a\\:description").text();
                    _ResultXML = null;
                    return _ResultXML;
                }

                return _ResultXML;
            },
            buildFetchRequest = function (fetch, RequestName) {
                alert("This method is obsolete and will be removed soon");
                var request = "<s:Envelope xmlns:s=\"http://schemas.xmlsoap.org/soap/envelope/\">";
                request += "<s:Body>";

                request += '<Execute xmlns="http://schemas.microsoft.com/xrm/2011/Contracts/Services">' +
                    '<request i:type="b:' + RequestName + 'Request" ' +
                    'xmlns:b="http://schemas.microsoft.com/xrm/2011/Contracts" ' +
                    'xmlns:i="http://www.w3.org/2001/XMLSchema-instance">' +
                    '<b:Parameters xmlns:c="http://schemas.datacontract.org/2004/07/System.Collections.Generic">' +
                    '<b:KeyValuePairOfstringanyType>' +
                    '<c:key>Query</c:key>' +
                    '<c:value i:type="b:FetchExpression">' +
                    '<b:Query>';

                request += CrmEncodeDecode.CrmXmlEncode(fetch);

                request += '</b:Query>' +
                    '</c:value>' +
                    '</b:KeyValuePairOfstringanyType>' +
                    '</b:Parameters>' +
                    '<b:RequestId i:nil="true"/>' +

                    '<b:RequestName>' + RequestName + '</b:RequestName>' +
                    '</request>' +
                    '</Execute>';

                request += '</s:Body></s:Envelope>';
                return request;

            },
            buildCustomActionRequest = function (entityId, entityName, requestName) {
                alert("This method is obsolete and will be removed soon");
                /// <summary>Creates request to execute custom action.</summary>
                /// <param name="entityId" type="Guid">Entity record id of type GUID.</param>
                /// <param name="entityName" type="String">Entity name of type String.</param>
                /// <param name="requestName" type="String">Custom Action name of type string.</param>
                /// <returns type="String">Returns soap xml request of type String.</returns>
                // Creating the request XML for calling the Action
                var requestXML = "";
                requestXML += "<s:Envelope xmlns:s=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
                    "<s:Body>" +
                    "<Execute xmlns=\"http://schemas.microsoft.com/xrm/2011/Contracts/Services\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                    "<request xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\">" +
                    "<a:Parameters xmlns:b=\"http://schemas.datacontract.org/2004/07/System.Collections.Generic\">" +
                    "<a:KeyValuePairOfstringanyType>" +
                    "<b:key>Target</b:key>" +
                    "<b:value i:type=\"a:EntityReference\">" +
                    "<a:Id>" + entityId + "</a:Id>" +
                    "<a:LogicalName>" + entityName + "</a:LogicalName>" +
                    "<a:Name i:nil=\"true\"/>" +
                    "</b:value>" +
                    "</a:KeyValuePairOfstringanyType>" +
                    "</a:Parameters>" +
                    "<a:RequestId i:nil=\"true\"/>" +
                    "<a:RequestName>" + requestName + "</a:RequestName>" +
                    "</request>" +
                    "</Execute>" +
                    "</s:Body>" +
                    "</s:Envelope>";
                return requestXML;
            },

            buildCustomActionRequestWithClientSideInputs = function (entityId, entityName, requestName, parameters) {
                alert("This method is obsolete and will be removed soon");
                /// <summary>Creates request to execute custom action.</summary>
                /// <param name="entityId" type="Guid">Entity record id of type GUID.</param>
                /// <param name="entityName" type="String">Entity name of type String.</param>
                /// <param name="requestName" type="String">Custom Action name of type string.</param>
                /// <returns type="String">Returns soap xml request of type String.</returns>
                // Creating the request XML for calling the Action
                var requestXML = "";
                requestXML += "<s:Envelope xmlns:s=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
                    "<s:Body>" +
                    "<Execute xmlns=\"http://schemas.microsoft.com/xrm/2011/Contracts/Services\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                    "<request xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\">" +
                    "<a:Parameters xmlns:b=\"http://schemas.datacontract.org/2004/07/System.Collections.Generic\">" +
                    "<a:KeyValuePairOfstringanyType>" +
                    "<b:key>Target</b:key>" +
                    "<b:value i:type=\"a:EntityReference\">" +
                    "<a:Id>" + entityId + "</a:Id>" +
                    "<a:LogicalName>" + entityName + "</a:LogicalName>" +
                    "<a:Name i:nil=\"true\"/>" +
                    "</b:value>" +
                    "</a:KeyValuePairOfstringanyType>";

                if (parameters !== null && typeof parameters !== "undefined") {

                    for (var key in parameters) {
                        if (!parameters.hasOwnProperty(key)) {
                            continue
                        } else {
                            requestXML += "<a:KeyValuePairOfstringanyType>";
                            requestXML += "<b:key>" + key + "</b:key>";
                            requestXML += "<b:value i:type=\"c:string\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">" + parameters[key] + "</b:value>";
                            requestXML += "</a:KeyValuePairOfstringanyType>";
                        }
                    }

                }

                requestXML += "</a:Parameters>" +
                    "<a:RequestId i:nil=\"true\"/>" +
                    "<a:RequestName>" + requestName + "</a:RequestName>" +
                    "</request>" +
                    "</Execute>" +
                    "</s:Body>" +
                    "</s:Envelope>";
                return requestXML;
            },
            businessEntity = function (logicalName, id) {
                ///<summary>
                /// A object represents a business entity for CRM 2011.
                ///</summary>
                ///<param name="logicalName" type="String">
                /// A String represents the name of the entity.
                /// For example, "contact" means the business entity will be a contact entity
                /// </param>
                ///<param name="id" type="String">
                /// A String represents the id of the entity. If not passed, it will be autopopulated as a empty guid string
                /// </param>
                this.id = (!id) ? "00000000-0000-0000-0000-000000000000" : id;
                this.logicalName = logicalName;
                this.attributes = new Object();
            };

        businessEntity.prototype = {

            /**
             * Deserialize an XML node into a CRM Business Entity object. The XML node comes from CRM Web Service's response.
             * @param {object} resultNode The XML node returned from CRM Web Service's Fetch, Retrieve, RetrieveMultiple messages.
             */
            deserialize: function (resultNode) {
                if (typeof jQuery == 'undefined') {
                    alert('jQuery is not loaded.\nPlease ensure that jQuery is included\n as webresource in the form load.');
                    return;
                }

                var obj = new Object();
                var resultNodes = resultNode.children();

                for (var j = 0; j < resultNodes.length; j++) {
                    var k;
                    var sKey;
                    switch (resultNodes.eq(j).prop("nodeName")) {
                        case "a:Attributes":
                            var attr = resultNodes[j];
                            for (k = 0; k < $(attr).children().length; k++) {
                                // Establish the Key for the Attribute
                                sKey = $(attr).children().eq(k).children(':first').text();
                                var sType = $(attr).children().eq(k).children().eq(1).attr('i:type');

                                var entRef;
                                var entCv;
                                switch (sType) {
                                    case "a:OptionSetValue":
                                        var entOsv = new xrmOptionSetValue();
                                        entOsv.type = sType.replace('a:', '');
                                        entOsv.value = parseInt($(attr).children().eq(k).children().eq(1).text());
                                        obj[sKey] = entOsv;
                                        break;

                                    case "a:EntityReference":
                                        entRef = new xrmEntityReference();
                                        entRef.type = sType.replace('a:', '');
                                        entRef.id = $(attr).children().eq(k).children().eq(1).children().eq(0).text();
                                        entRef.logicalName = $(attr).children().eq(k).children().eq(1).children().eq(1).text();
                                        entRef.name = $(attr).children().eq(k).children().eq(1).children().eq(2).text();
                                        obj[sKey] = entRef;
                                        break;

                                    case "a:EntityCollection":
                                        entRef = new xrmEntityCollection();
                                        entRef.type = sType.replace('a:', '');

                                        //get all party items....
                                        var items = [];
                                        for (var y = 0; y < $(attr).children().eq(k).children().eq(1).children().eq(0).children().length; y++) {
                                            var itemNodes = $(attr).children().eq(k).children().eq(1).children().eq(0).children().eq(y).children().eq(0).children();
                                            for (var z = 0; z < itemNodes.length; z++) {
                                                if (itemNodes.eq(z).children().eq(0).text() == "partyid") {
                                                    var itemRef = new xrmEntityReference();
                                                    itemRef.id = itemNodes.eq(z).children().eq(1).children().eq(0).text();
                                                    itemRef.logicalName = itemNodes.eq(z).children().eq(1).children().eq(1).text();
                                                    itemRef.name = itemNodes.eq(z).children().eq(1).children().eq(2).text();
                                                    items[y] = itemRef;
                                                }
                                            }
                                        }
                                        entRef.value = items;
                                        obj[sKey] = entRef;
                                        break;

                                    case "a:Money":
                                        entCv = new xrmValue();
                                        entCv.type = sType.replace('a:', '');
                                        entCv.value = parseFloat($(attr).children().eq(k).children().eq(1).text());
                                        obj[sKey] = entCv;
                                        break;

                                    default:
                                        entCv = new xrmValue();
                                        entCv.type = sType.replace('c:', '').replace('a:', '');
                                        if (entCv.type == "int") {
                                            entCv.value = parseInt($(attr).children().eq(k).children().eq(1).text());
                                        } else if (entCv.type == "decimal") {
                                            entCv.value = parseFloat($(attr).children().eq(k).children().eq(1).text());
                                        } else if (entCv.type == "dateTime") {
                                            entCv.value = Date.parse($(attr).children().eq(k).children().eq(1).text());
                                        } else if (entCv.type == "boolean") {
                                            entCv.value = ($(attr).children().eq(k).children().eq(1).text() == 'false') ? false : true;
                                        } else {
                                            entCv.value = $(attr).children().eq(k).children().eq(1).text();
                                        }
                                        obj[sKey] = entCv;
                                        break;
                                }
                            }
                            this.attributes = obj;
                            break;

                        case "a:Id":
                            this.id = $(resultNodes).eq(j).text();
                            break;

                        case "a:LogicalName":
                            this.logicalName = $(resultNodes).eq(j).text();
                            break;

                        case "a:FormattedValues":
                            var foVal = $(resultNodes).eq(j);

                            for (k = 0; k < foVal.children().length; k++) {
                                // Establish the Key, we are going to fill in the formatted value of the already found attribute
                                sKey = foVal.children().eq(k).children().eq(0).text();
                                this.attributes[sKey].formattedValue = foVal.children().eq(k).children().eq(1).text();
                                if (isNaN(this.attributes[sKey].value) && this.attributes[sKey].type == "dateTime") {
                                    this.attributes[sKey].value = new Date(this.attributes[sKey].formattedValue);
                                }
                            }
                            break;
                    }
                }
            },
        };
        return {
            ExecuteRequest: executeRequest,
            buildFetchRequest: buildFetchRequest,
            BuildCustomActionRequest: buildCustomActionRequest,
            BuildCustomActionRequestWithClientSideInputs: buildCustomActionRequestWithClientSideInputs,
            Fetch: fetch,
            RetrieveEntityMetadata: retrieveEntityMetadata
        };
    }());

    TnXrmUtilities.WebAPI = (function () {
        "use strict";
        //Private Properties
        //Start

        ///<summary>
        /// Private function to the context object.
        ///</summary>
        ///<returns>Context</returns>
        var context = function () {

            if (typeof (GetGlobalContext) !== "undefined") {
                return GetGlobalContext();
            }

            if (Xrm !== undefined && Xrm !== null) {
                return Xrm.Page.context;
            }
            throw new Error("Context is not available.");
        },

            ///<summary>
            /// Private function to return the server URL from the context
            ///</summary>
            ///<returns>String</returns>
            getClientUrl = function () {
                var clientUrl = context().getClientUrl();
                return clientUrl;
            },

            ///<summary>
            /// Private function to return the path to the REST endpoint.
            ///</summary>
            ///<returns>String</returns>
            getWebAPIPath = function () {
                return getClientUrl() + "/api/data/v8.0/";
            },

            ///<summary>
            /// Private function return an Error object to the errorCallback
            ///</summary>
            ///<param name="req" type="XMLHttpRequest">
            /// The XMLHttpRequest response that returned an error.
            ///</param>
            ///<returns>Error</returns>
            errorHandler = function (req) {
                //Error descriptions come from http://support.microsoft.com/kb/193625
                switch (req.status) {
                    case 12029:
                        return new Error("The attempt to connect to the server failed.");
                        break;
                    case 12007:
                        return new Error("The server name could not be resolved.");
                        break;
                    case 503:
                        return new Error(req.statusText + " Status Code:" + req.status + " The Web API Preview is not enabled.");
                        break;
                    default:
                        var errorText;
                        try {
                            var errorObj = JSON.parse(req.responseText).error;
                            errorText = JSON.parse(req.responseText).error.message;
                            if (errorObj.innererror !== undefined && errorObj.innererror !== null && errorText !== errorObj.innererror.message) {
                                errorText += "\nInner Message: " + errorObj.innererror.message +
                                    "\nType: " + errorObj.innererror.type;
                            }
                        } catch (e) {
                            errorText = req.responseText;
                        }
                        return new Error("Error : " +
                            req.status + ": " +
                            req.statusText + ": " + errorText);
                        break;

                }
            },

            ///<summary>
            /// Private function to convert matching string values to Date objects.
            ///</summary>
            ///<param name="key" type="String">
            /// The key used to identify the object property
            ///</param>
            ///<param name="value" type="String">
            /// The string value representing a date
            ///</param>
            dateReviver = function (key, value) {
                var a;
                if (typeof value === 'string') {
                    a = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(?:\.\d*)?)Z$/.exec(value);
                    if (a) {
                        return new Date(Date.UTC(+a[1], +a[2] - 1, +a[3], +a[4], +a[5], +a[6]));
                    }
                }
                return value;
            },

            ///<summary>
            /// Private function used to check whether required parameters are null or undefined
            ///</summary>
            ///<param name="parameter" type="Object">
            /// The parameter to check;
            ///</param>
            ///<param name="message" type="String">
            /// The error message text to include when the error is thrown.
            ///</param>
            parameterCheck = function (parameter, message) {
                if ((parameter === undefined) || parameter === null) {
                    throw new Error(message);
                }
            },

            ///<summary>
            /// Private function used to check whether required parameters are null or undefined
            ///</summary>
            ///<param name="parameter" type="String">
            /// The string parameter to check;
            ///</param>
            ///<param name="message" type="String">
            /// The error message text to include when the error is thrown.
            ///</param>
            stringParameterCheck = function (parameter, message) {
                if (typeof parameter !== "string") {
                    throw new Error(message);
                }
            },

            ///<summary>
            /// Private function used to check whether required callback parameters are functions
            ///</summary>
            ///<param name="callbackParameter" type="Function">
            /// The callback parameter to check;
            ///</param>
            ///<param name="message" type="String">
            /// The error message text to include when the error is thrown.
            ///</param>
            callbackParameterCheck = function (callbackParameter, message) {
                if (typeof callbackParameter !== "function") {
                    throw new Error(message);
                }
            },

            ///<summary>
            /// Formats the lookup attribute name
            ///</summary>
            formatLookupAttributeName = function (attributeName) {
                return "_" + attributeName + "_value";
            },

            ///<summary>
            /// Gets the lookup attribute value
            ///</summary>
            getLookupAttributeValue = function (result, attributeName) {
                if (result !== undefined && result !== null) {
                    if (result["_" + attributeName + "_value"] !== undefined && result["_" + attributeName + "_value"] !== null) {
                        var lookupId = result["_" + attributeName + "_value"];
                        if (lookupId.indexOf('{') == -1)
                            lookupId = '{' + lookupId;
                        if (lookupId.indexOf('}') == -1)
                            lookupId = lookupId + '}';
                        lookupId = lookupId.toUpperCase();
                        return {
                            Id: lookupId,
                            LogicalName: "",
                            Name: (result["_" + attributeName + "_value@OData.Community.Display.V1.FormattedValue"] !== undefined && result["_" + attributeName + "_value@OData.Community.Display.V1.FormattedValue"] !== null) ? result["_" + attributeName + "_value@OData.Community.Display.V1.FormattedValue"] : null
                        };
                    }
                }

                return null;
            },

            ///<summary>
            /// Converts the Date to EDM format
            ///<param name="yy" type="string">
            ///FullYear(YYYY)
            ///</param>
            ///<param name="mm" type="string">
            ///month(MM)
            ///</param>
            ///<param name="dd" type="string">
            ///Date(DD)
            ///</param>
            ///<param name="hh" type="string">
            ///Hours(HH)
            ///</param>
            ///<param name="min" type="string">
            ///Minusts(MM)
            ///</param>
            ///<param name="ss" type="string">
            ///Seconds(SS)
            ///</param>
            ///</summary>
            formatDateEDM = function (yy, mm, dd, hh, min, ss) {
                if (yy !== undefined && yy !== null && mm !== undefined && mm !== null && dd !== undefined && dd !== null) {
                    if (mm.toString().length < 2) {
                        mm = '0' + mm;
                    }
                    if (dd.toString().length < 2) {
                        dd = '0' + dd;
                    }
                    if (hh !== undefined && hh !== null && min !== undefined && min !== null && ss !== undefined && ss !== null) {
                        if (hh.toString().length < 2) {
                            hh = '0' + hh;
                        }
                        if (min.toString().length < 2) {
                            min = '0' + min;
                        }
                        if (ss.toString().length < 2) {
                            ss = '0' + ss;
                        }
                        return [yy, mm, dd].join('-') + 'T' + [hh, min, ss].join(':') + 'Z';
                    } else {
                        return [yy, mm, dd].join('-') + 'T' + ["00", "00", "00"].join(':') + 'Z';
                    }
                }

                return null;
            },

            ///<summary>
            /// Sets the lookup attribute value for update operations
            ///</summary>
            setLookupAttributeValue = function (object, attributeName, entityName, entityId) {
                if (entityName === undefined || entityName === null) {
                    object[attributeName + "@odata.bind"] = null;
                    return;
                }

                if (entityId !== undefined && entityId !== null) {
                    entityId = entityId.replace("{", "").replace("}", "");
                }

                object[attributeName + "@odata.bind"] = "/" + entityName + "(" + entityId + ")";
            },

            setRequestHeaders = function (req, includeFormattedValues, includeLookupProperties) {
                req.setRequestHeader("Accept", "application/json");
                req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                req.setRequestHeader("OData-MaxVersion", "4.0");
                req.setRequestHeader("OData-Version", "4.0");
                if (!(includeLookupProperties !== undefined && includeLookupProperties !== null && includeLookupProperties === false)) {
                    req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"")
                }
                else if (!(includeFormattedValues !== undefined && includeFormattedValues !== null && includeFormattedValues === false)) {
                    req.setRequestHeader("Prefer", "odata.include-annotations=\"OData.Community.Display.V1.FormattedValue\"")
                }
            },

            ///<summary>
            /// Sends an asynchronous request to create a new record.
            ///</summary>
            ///<param name="object" type="Object">
            /// A JavaScript object with properties corresponding to the Schema name of
            /// entity attributes that are valid for create operations.
            ///</param>
            ///<param name="type" type="String">
            /// The Schema Name of the Entity type record to create.
            /// For an Account record, use "Account"
            ///</param>
            ///<param name="successCallback" type="Function">
            /// The function that will be passed through and be called by a successful response. 
            /// This function can accept the returned record as a parameter.
            /// </param>
            ///<param name="errorCallback" type="Function">
            /// The function that will be passed through and be called by a failed response. 
            /// This function must accept an Error object as a parameter.
            /// </param>
            createRecord = function (object, type, successCallback, errorCallback) {
                parameterCheck(object, "TnXrmUtilities.WebAPI.createRecord requires the object parameter.");
                stringParameterCheck(type, "TnXrmUtilities.WebAPI.createRecord requires the type parameter is a string.");
                callbackParameterCheck(successCallback, "TnXrmUtilities.WebAPI.createRecord requires the successCallback is a function.");
                callbackParameterCheck(errorCallback, "TnXrmUtilities.WebAPI.createRecord requires the errorCallback is a function.");

                var req = new XMLHttpRequest();
                req.open("POST", encodeURI(getWebAPIPath() + type), true);
                setRequestHeaders(req, false);

                req.onreadystatechange = function () {
                    /* complete */
                    if (this.readyState === 4) {
                        req.onreadystatechange = null;
                        if (this.status === 204) {
                            var responseUrl = this.getResponseHeader("OData-EntityId"),
                                entityId;
                            if (responseUrl !== undefined && responseUrl !== null) {
                                entityId = responseUrl.substr(responseUrl.length - 38).substring(1, 37);
                            }
                            successCallback(entityId);
                        } else {
                            errorCallback(TnXrmUtilities.WebAPI.ErrorHandler(this));
                        }
                    }
                };
                req.send(JSON.stringify(object));
            },

            ///<summary>
            /// Sends an synchronous request to create a new record.
            ///</summary>
            ///<param name="object" type="Object">
            /// A JavaScript object with properties corresponding to the Schema name of
            /// entity attributes that are valid for create operations.
            ///</param>
            ///<param name="type" type="String">
            /// The Schema Name of the Entity type record to create.
            /// For an Account record, use "Account"
            ///</param>
            ///<param name="successCallback" type="Function">
            /// The function that will be passed through and be called by a successful response. 
            /// This function can accept the returned record as a parameter.
            /// </param>
            ///<param name="errorCallback" type="Function">
            /// The function that will be passed through and be called by a failed response. 
            /// This function must accept an Error object as a parameter.
            /// </param>
            createRecordSync = function (object, type, successCallback, errorCallback) {
                parameterCheck(object, "TnXrmUtilities.WebAPI.createRecord requires the object parameter.");
                stringParameterCheck(type, "TnXrmUtilities.WebAPI.createRecord requires the type parameter is a string.");
                callbackParameterCheck(successCallback, "TnXrmUtilities.WebAPI.createRecord requires the successCallback is a function.");
                callbackParameterCheck(errorCallback, "TnXrmUtilities.WebAPI.createRecord requires the errorCallback is a function.");

                var req = new XMLHttpRequest();
                req.open("POST", encodeURI(getWebAPIPath() + type), false);
                setRequestHeaders(req, false);

                req.onreadystatechange = function () {
                    /* complete */
                    if (this.readyState === 4) {
                        req.onreadystatechange = null;
                        if (this.status === 204) {
                            var responseUrl = this.getResponseHeader("OData-EntityId"),
                                entityId;
                            if (responseUrl !== undefined && responseUrl !== null) {
                                entityId = responseUrl.substr(responseUrl.length - 38).substring(1, 37);
                            }
                            successCallback(entityId);
                        } else {
                            errorCallback(TnXrmUtilities.WebAPI.ErrorHandler(this));
                        }
                    }
                };
                req.send(JSON.stringify(object));
            },

            ///<summary>
            /// Sends an asynchronous request to retrieve a record.
            ///</summary>
            ///<param name="id" type="String">
            /// A String representing the GUID value for the record to retrieve.
            ///</param>
            ///<param name="type" type="String">
            /// The Schema Name of the Entity type record to retrieve.
            /// For an Account record, use "Account"
            ///</param>
            ///<param name="select" type="String">
            /// A String representing the $select OData System Query Option to control which
            /// attributes will be returned. This is a comma separated list of Attribute names that are valid for retrieve.
            /// If null all properties for the record will be returned
            ///</param>
            ///<param name="expand" type="String">
            /// A String representing the $expand OData System Query Option value to control which
            /// related records are also returned. This is a comma separated list of of up to 6 entity relationship names
            /// If null no expanded related records will be returned.
            ///</param>
            ///<param name="successCallback" type="Function">
            /// The function that will be passed through and be called by a successful response. 
            /// This function must accept the returned record as a parameter.
            /// </param>
            ///<param name="errorCallback" type="Function">
            /// The function that will be passed through and be called by a failed response. 
            /// This function must accept an Error object as a parameter.
            /// </param>
            retrieveRecord = function (id, type, select, successCallback, errorCallback) {
                stringParameterCheck(id, "TnXrmUtilities.WebAPI.retrieveRecord requires the id parameter is a string.");
                id = id.replace("{", "").replace("}", "");
                stringParameterCheck(type, "TnXrmUtilities.WebAPI.retrieveRecord requires the type parameter is a string.");
                if (select !== undefined && select !== null) {
                    stringParameterCheck(select, "TnXrmUtilities.WebAPI.retrieveRecord requires the select parameter is a string.");
                }
                callbackParameterCheck(successCallback, "TnXrmUtilities.WebAPI.retrieveRecord requires the successCallback parameter is a function.");
                callbackParameterCheck(errorCallback, "TnXrmUtilities.WebAPI.retrieveRecord requires the errorCallback parameter is a function.");

                var req, systemQueryOptions = "";

                if (select !== null) {
                    systemQueryOptions = "?$select=" + select;
                }

                req = new XMLHttpRequest();
                req.open("GET", getWebAPIPath() + type + "(" + id + ")" + systemQueryOptions, true);
                setRequestHeaders(req);

                req.onreadystatechange = function () {
                    /* complete */
                    if (this.readyState === 4) {
                        req.onreadystatechange = null;
                        if (this.status === 200) {
                            if (Object.keys(JSON.parse(this.responseText, TnXrmUtilities.WebAPI.dateReviver)).length > 0) {
                                successCallback(JSON.parse(this.responseText, dateReviver));
                            } else {
                                successCallback(null);
                            }
                        } else {
                            errorCallback(TnXrmUtilities.WebAPI.ErrorHandler(this));
                        }
                    }
                };
                req.send();
            },


            ///<summary>
            /// Sends an synchronous request to retrieve a record.
            ///</summary>
            ///<param name="id" type="String">
            /// A String representing the GUID value for the record to retrieve.
            ///</param>
            ///<param name="type" type="String">
            /// The Schema Name of the Entity type record to retrieve.
            /// For an Account record, use "Account"
            ///</param>
            ///<param name="select" type="String">
            /// A String representing the $select OData System Query Option to control which
            /// attributes will be returned. This is a comma separated list of Attribute names that are valid for retrieve.
            /// If null all properties for the record will be returned
            ///</param>
            retrieveRecordSync = function (id, type, select) {
                stringParameterCheck(id, "TnXrmUtilities.WebAPI.retrieveRecord requires the id parameter is a string.");
                id = id.replace("{", "").replace("}", "");
                stringParameterCheck(type, "TnXrmUtilities.WebAPI.retrieveRecord requires the type parameter is a string.");
                if (select !== undefined && select !== null) {
                    stringParameterCheck(select, "TnXrmUtilities.WebAPI.retrieveRecord requires the select parameter is a string.");
                }

                var req, results, systemQueryOptions = "";

                if (select !== null) {
                    systemQueryOptions = "?$select=" + select;
                }

                req = new XMLHttpRequest();
                req.open("GET", getWebAPIPath() + type + "(" + id + ")" + systemQueryOptions, false);
                setRequestHeaders(req);

                req.onreadystatechange = function () {
                    /* complete */
                    if (this.readyState === 4) {
                        req.onreadystatechange = null;
                        if (this.status === 200) {

                            if (Object.keys(JSON.parse(this.responseText, TnXrmUtilities.WebAPI.dateReviver)).length > 0) {
                                results = JSON.parse(this.responseText, TnXrmUtilities.WebAPI.dateReviver);
                            } else {
                                results = null;
                            }

                        } else {
                            alert(TnXrmUtilities.WebAPI.ErrorHandler(this));
                        }
                    }
                };
                req.send();

                return results;
            },

            ///<summary>
            /// Sends an asynchronous request to update a record.
            ///</summary>
            ///<param name="id" type="String">
            /// A String representing the GUID value for the record to retrieve.
            ///</param>
            ///<param name="object" type="Object">
            /// A JavaScript object with properties corresponding to the Schema Names for
            /// entity attributes that are valid for update operations.
            ///</param>
            ///<param name="type" type="String">
            /// The Schema Name of the Entity type record to retrieve.
            /// For an Account record, use "Account"
            ///</param>
            ///<param name="successCallback" type="Function">
            /// The function that will be passed through and be called by a successful response. 
            /// Nothing will be returned to this function.
            /// </param>
            ///<param name="errorCallback" type="Function">
            /// The function that will be passed through and be called by a failed response. 
            /// This function must accept an Error object as a parameter.
            /// </param>
            updateRecord = function (id, object, type, successCallback, errorCallback) {
                stringParameterCheck(id, "TnXrmUtilities.WebAPI.updateRecord requires the id parameter.");
                id = id.replace("{", "").replace("}", "");
                parameterCheck(object, "TnXrmUtilities.WebAPI.updateRecord requires the object parameter.");
                stringParameterCheck(type, "TnXrmUtilities.WebAPI.updateRecord requires the type parameter.");
                callbackParameterCheck(successCallback, "TnXrmUtilities.WebAPI.updateRecord requires the successCallback is a function.");
                callbackParameterCheck(errorCallback, "TnXrmUtilities.WebAPI.updateRecord requires the errorCallback is a function.");

                var propertyCount = 0;

                for (var property in object) {
                    propertyCount++;
                    if (object[property] === null) {
                        var deleteResponse = deleteSingleAttributeFromRecord(id, type, property);
                        if (deleteResponse !== undefined && deleteResponse !== null) {
                            if (!deleteResponse.status) {
                                errorCallback(deleteResponse.error);
                                return;
                            } else {
                                delete object[property];
                                propertyCount--;
                            }
                        }
                    }
                }

                if (propertyCount > 0) {
                    var req = new XMLHttpRequest();
                    req.open("PATCH", encodeURI(getWebAPIPath() + type + "(" + id + ")"), true);
                    setRequestHeaders(req, false);

                    req.onreadystatechange = function () {
                        /* complete */
                        if (this.readyState === 4) {
                            req.onreadystatechange = null;
                            if (this.status === 204 || this.status === 1223) {
                                successCallback();
                            } else {
                                errorCallback(TnXrmUtilities.WebAPI.ErrorHandler(this));
                            }
                        }
                    };
                    req.send(JSON.stringify(object));
                } else {
                    successCallback();
                }
            },

             ///<summary>
            /// Sends a synchronous request to update a record.
            ///</summary>
            ///<param name="id" type="String">
            /// A String representing the GUID value for the record to retrieve.
            ///</param>
            ///<param name="object" type="Object">
            /// A JavaScript object with properties corresponding to the Schema Names for
            /// entity attributes that are valid for update operations.
            ///</param>
            ///<param name="type" type="String">
            /// The Schema Name of the Entity type record to retrieve.
            /// For an Account record, use "Account"
            ///</param>
            ///<param name="successCallback" type="Function">
            /// The function that will be passed through and be called by a successful response. 
            /// Nothing will be returned to this function.
            /// </param>
            ///<param name="errorCallback" type="Function">
            /// The function that will be passed through and be called by a failed response. 
            /// This function must accept an Error object as a parameter.
            /// </param>
            updateRecordSync = function (id, object, type, successCallback, errorCallback) {
                stringParameterCheck(id, "TnXrmUtilities.WebAPI.updateRecordSync requires the id parameter.");
                id = id.replace("{", "").replace("}", "");
                parameterCheck(object, "TnXrmUtilities.WebAPI.updateRecordSync requires the object parameter.");
                stringParameterCheck(type, "TnXrmUtilities.WebAPI.updateRecordSync requires the type parameter.");
                callbackParameterCheck(successCallback, "TnXrmUtilities.WebAPI.updateRecordSync requires the successCallback is a function.");
                callbackParameterCheck(errorCallback, "TnXrmUtilities.WebAPI.updateRecordSync requires the errorCallback is a function.");

                var propertyCount = 0;

                for (var property in object) {
                    propertyCount++;
                    if (object[property] === null) {
                        var deleteResponse = deleteSingleAttributeFromRecord(id, type, property);
                        if (deleteResponse !== undefined && deleteResponse !== null) {
                            if (!deleteResponse.status) {
                                errorCallback(deleteResponse.error);
                                return;
                            } else {
                                delete object[property];
                                propertyCount--;
                            }
                        }
                    }
                }

                if (propertyCount > 0) {
                    var req = new XMLHttpRequest();
                    req.open("PATCH", encodeURI(getWebAPIPath() + type + "(" + id + ")"), false);
                    setRequestHeaders(req, false);

                    req.onreadystatechange = function () {
                        /* complete */
                        if (this.readyState === 4) {
                            req.onreadystatechange = null;
                            if (this.status === 204 || this.status === 1223) {
                                successCallback();
                            } else {
                                errorCallback(TnXrmUtilities.WebAPI.ErrorHandler(this));
                            }
                        }
                    };
                    req.send(JSON.stringify(object));
                } else {
                    successCallback();
                }
            },

            ///<summary>
            /// Sends an asynchronous request to delete a record.
            ///</summary>
            ///<param name="id" type="String">
            /// A String representing the GUID value for the record to delete.
            ///</param>
            ///<param name="type" type="String">
            /// The Schema Name of the Entity type record to delete.
            /// For an Account record, use "Account"
            ///</param>
            ///<param name="successCallback" type="Function">
            /// The function that will be passed through and be called by a successful response. 
            /// Nothing will be returned to this function.
            /// </param>
            ///<param name="errorCallback" type="Function">
            /// The function that will be passed through and be called by a failed response. 
            /// This function must accept an Error object as a parameter.
            /// </param>
            deleteRecord = function (id, type, successCallback, errorCallback) {
                stringParameterCheck(id, "TnXrmUtilities.WebAPI.deleteRecord requires the id parameter.");
                id = id.replace("{", "").replace("}", "");
                stringParameterCheck(type, "TnXrmUtilities.WebAPI.deleteRecord requires the type parameter.");
                callbackParameterCheck(successCallback, "TnXrmUtilities.WebAPI.deleteRecord requires the successCallback is a function.");
                callbackParameterCheck(errorCallback, "TnXrmUtilities.WebAPI.deleteRecord requires the errorCallback is a function.");

                var req = new XMLHttpRequest();
                req.open("DELETE", encodeURI(getWebAPIPath() + type + "(" + id + ")"), true);
                setRequestHeaders(req, false);

                req.onreadystatechange = function () {
                    /* complete */
                    if (this.readyState === 4) {
                        req.onreadystatechange = null;
                        if (this.status === 204 || this.status === 1223) {
                            successCallback();
                        } else {
                            errorCallback(TnXrmUtilities.WebAPI.ErrorHandler(this));
                        }
                    }
                };
                req.send();
            },

            deleteSingleAttributeFromRecord = function (id, type, attributeName) {
                var result = {};
                stringParameterCheck(id, "TnXrmUtilities.WebAPI.deleteSingleAttributeFromRecord requires the id parameter.");
                id = id.replace("{", "").replace("}", "");
                stringParameterCheck(type, "TnXrmUtilities.WebAPI.deleteSingleAttributeFromRecord requires the type parameter.");

                if (attributeName.indexOf("@odata.bind") > -1) {
                    attributeName = attributeName.replace("@odata.bind", "") + "/$ref";
                }

                var req = new XMLHttpRequest();
                req.open("DELETE", encodeURI(getWebAPIPath() + type + "(" + id + ")/" + attributeName), false);
                setRequestHeaders(req, false);

                req.onreadystatechange = function () {
                    /* complete */
                    if (this.readyState === 4) {
                        req.onreadystatechange = null;
                        if (this.status === 204 || this.status === 1223) {
                            result.status = true;
                        } else {
                            result.status = false;
                            result.error = TnXrmUtilities.WebAPI.ErrorHandler(this);
                        }
                    }
                };
                req.send();

                return result;
            },

            ///<summary>
            /// Sends an asynchronous request to retrieve records.
            ///</summary>
            ///<param name="type" type="String">
            /// The Schema Name of the Entity type record to retrieve.
            /// For an Account record, use "Account"
            ///</param>
            ///<param name="options" type="String">
            /// A String representing the OData System Query Options to control the data returned
            ///</param>
            ///<param name="successCallback" type="Function">
            /// The function that will be passed through and be called for each page of records returned.
            /// Each page is 50 records. If you expect that more than one page of records will be returned,
            /// this function should loop through the results and push the records into an array outside of the function.
            /// Use the OnComplete event handler to know when all the records have been processed.
            /// </param>
            ///<param name="errorCallback" type="Function">
            /// The function that will be passed through and be called by a failed response. 
            /// This function must accept an Error object as a parameter.
            /// </param>
            ///<param name="OnComplete" type="Function">
            /// The function that will be called when all the requested records have been returned.
            /// No parameters are passed to this function.
            /// </param>
            retrieveMultipleRecords = function (type, options, successCallback, errorCallback, OnComplete) {
                stringParameterCheck(type, "TnXrmUtilities.WebAPI.retrieveMultipleRecords requires the type parameter is a string.");
                if (options !== undefined && options !== null) {
                    stringParameterCheck(options, "TnXrmUtilities.WebAPI.retrieveMultipleRecords requires the options parameter is a string.");
                }
                callbackParameterCheck(successCallback, "TnXrmUtilities.WebAPI.retrieveMultipleRecords requires the successCallback parameter is a function.");
                callbackParameterCheck(errorCallback, "TnXrmUtilities.WebAPI.retrieveMultipleRecords requires the errorCallback parameter is a function.");
                callbackParameterCheck(OnComplete, "TnXrmUtilities.WebAPI.retrieveMultipleRecords requires the OnComplete parameter is a function.");

                var optionsString, req;
                if (options !== null) {
                    if (options.charAt(0) !== "?") {
                        optionsString = "?" + options;
                    } else {
                        optionsString = options;
                    }
                }

                req = new XMLHttpRequest();
                req.open("GET", getWebAPIPath() + type + optionsString, true);
                setRequestHeaders(req);

                req.onreadystatechange = function () {
                    /* complete */
                    if (this.readyState === 4) {
                        req.onreadystatechange = null;
                        switch (this.status) {
                            case 200:
                                var queryOptions,
                                    returned = JSON.parse(this.responseText, dateReviver);

                                // Invoke the Callback method  
                                successCallback(returned.value);

                                if (returned['@odata.nextLink'] !== undefined && returned['@odata.nextLink'] !== null) {
                                    queryOptions = returned['@odata.nextLink'].substring((getWebAPIPath() + type).length);
                                    retrieveMultipleRecords(type, queryOptions, successCallback, errorCallback, OnComplete);
                                } else {
                                    OnComplete();
                                }
                                break;
                            default:
                                errorCallback(TnXrmUtilities.WebAPI.ErrorHandler(this));
                                break;
                        }
                    }
                };
                req.send();
            },
            retrieveMultipleRecordsSync = function (type, options, errorCallback) {
                stringParameterCheck(type, "TnXrmUtilities.WebAPI.retrieveMultipleRecordsSync requires the type parameter is a string.");
                if (options !== undefined && options !== null) {
                    stringParameterCheck(options, "TnXrmUtilities.WebAPI.retrieveMultipleRecordsSync requires the options parameter is a string.");
                }

                var optionsString, req, results;
                if (options !== null) {
                    if (options.charAt(0) !== "?") {
                        optionsString = "?" + options;
                    } else {
                        optionsString = options;
                    }
                }

                req = new XMLHttpRequest();
                req.open("GET", getWebAPIPath() + type + optionsString, false);
                setRequestHeaders(req);

                req.onreadystatechange = function () {
                    /* complete */
                    if (this.readyState === 4) {
                        req.onreadystatechange = null;
                        switch (this.status) {
                            case 200:
                                var queryOptions,
                                    returned = JSON.parse(this.responseText, dateReviver);
                                results = returned.value;
                                break;
                            default:
                                if (errorCallback !== undefined && errorCallback !== null) {
                                    errorCallback(TnXrmUtilities.WebAPI.ErrorHandler(this));
                                }
                                break;
                        }
                    }
                };
                req.send();

                return results;
            },

            retrieveRecordObjectSync = function (type, options, errorCallback) {
                stringParameterCheck(type, "TnXrmUtilities.WebAPI.retrieveMultipleRecordsSync requires the type parameter is a string.");
                if (options !== undefined && options !== null) {
                    stringParameterCheck(options, "TnXrmUtilities.WebAPI.retrieveMultipleRecordsSync requires the options parameter is a string.");
                }

                var optionsString, req, results;
                if (options !== null) {
                    if (options.charAt(0) !== "?") {
                        optionsString = "?" + options;
                    } else {
                        optionsString = options;
                    }
                }

                req = new XMLHttpRequest();
                req.open("GET", getWebAPIPath() + type + optionsString, false);
                setRequestHeaders(req);

                req.onreadystatechange = function () {
                    /* complete */
                    if (this.readyState === 4) {
                        req.onreadystatechange = null;
                        switch (this.status) {
                            case 200:
                                var queryOptions,
                                    returned = JSON.parse(this.responseText, dateReviver);
                                if (returned.value !== undefined) {
                                    results = returned.value;
                                } else {
                                    results = returned;
                                }
                                break;
                            default:
                                if (errorCallback !== undefined && errorCallback !== null) {
                                    errorCallback(TnXrmUtilities.WebAPI.ErrorHandler(this));
                                }
                                break;
                        }
                    }
                };
                req.send();

                return results;
            },

            executeCustomActionSync = function (entityId, entityCollectionName, actionName, parameters, isGlobleAction) {
                var query = "",
                    req = "",
                    response = {};
                stringParameterCheck(entityId, "TnXrmUtilities.WebAPI.executeCustomActionSync requires the e entityId parameter is a string.");
                stringParameterCheck(entityCollectionName, "TnXrmUtilities.WebAPI.executeCustomActionSync requires the entityCollectionName parameter is a string.");
                stringParameterCheck(actionName, "TnXrmUtilities.WebAPI.executeCustomActionSync requires the actionName parameter is a string.");
                entityId = entityId.replace("{", "").replace("}", "");

                if (isGlobleAction !== undefined && isGlobleAction !== null && isGlobleAction !== true) {
                    query = entityCollectionName + "(" + entityId + ")/Microsoft.Dynamics.CRM." + actionName;
                } else {
                    query = entityCollectionName + "(" + entityId + ")/" + actionName;
                }

                req = new XMLHttpRequest();
                req.open("POST", getWebAPIPath() + query, false);
                setRequestHeaders(req, false);
                req.onreadystatechange = function () {
                    if (this.readyState == 4) {
                        req.onreadystatechange = null;
                        switch (this.status) {
                            case 200:
                                response.status = true;
                                response.results = JSON.parse(this.response);
                                break;
                            case 204:
                                response.status = true;
                                break;
                            case 1223:
                                response.status = true;
                                break;
                            default:
                                response.status = false;
                                response.error = JSON.parse(this.response).error;
                                break;
                        }

                    }
                };

                req.send((parameters !== undefined && parameters !== null) ? JSON.stringify(parameters) : null);
                return response;
            },

            executeCustomGlobalActionSync = function (parameters, actionName) {
                var results = null;
                var req = new XMLHttpRequest();
                req.open("POST", Xrm.Page.context.getClientUrl() + "/api/data/v8.1/" + actionName, false);
                req.setRequestHeader("OData-MaxVersion", "4.0");
                req.setRequestHeader("OData-Version", "4.0");
                req.setRequestHeader("Accept", "application/json");
                req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                req.onreadystatechange = function () {
                    if (this.readyState === 4) {
                        req.onreadystatechange = null;
                        if (this.status === 200) {
                            results = JSON.parse(this.response);
                        } else {
                            Xrm.Utility.alertDialog(this.statusText);
                        }
                    }
                };
                req.send(JSON.stringify(parameters));
                return results;
            },

            executeFetchXML = function (type, fetchXml, errorCallback) {
                'use strict';
                stringParameterCheck(type, "TnXrmUtilities.WebAPI.ExecuteFetchXml requires the type parameter is a string.");
                stringParameterCheck(fetchXml, "TnXrmUtilities.WebAPI.ExecuteFetchXml requires the fetchXml parameter is a string.");
                var fetch, req, results = [];
                fetch = escape(fetchXml);
                req = new XMLHttpRequest();
                req.open("GET", getWebAPIPath() + type + "?fetchXml=" + fetch, false);
                setRequestHeaders(req);
                req.onreadystatechange = function () {
                    /* complete */
                    if (this.readyState === 4) {
                        req.onreadystatechange = null;
                        switch (this.status) {
                            case 200:
                                var result = this.response;
                                results = JSON.parse(result).value;
                                break;
                            default:
                                if (errorCallback !== 'undefined' && typeof errorCallback === 'function') {
                                    errorCallback(TnXrmUtilities.WebAPI.ErrorHandler(this));
                                }
                                break;
                        }
                    }
                };
                req.send();
                return results;
            };


        // Toolkit's public static members
        return {
            ErrorHandler: errorHandler,
            FormatLookupAttributeName: formatLookupAttributeName,
            GetLookupAttributeValue: getLookupAttributeValue,
            SetLookupAttributeValue: setLookupAttributeValue,
            CreateRecord: createRecord,
            CreateRecordSync: createRecordSync,
            RetrieveRecord: retrieveRecord,
            RetrieveRecordSync: retrieveRecordSync,
            UpdateRecord: updateRecord,
            UpdateRecordSync: updateRecordSync,
            DeleteRecord: deleteRecord,
            RetrieveMultipleRecords: retrieveMultipleRecords,
            RetrieveMultipleRecordsSync: retrieveMultipleRecordsSync,
            RetrieveRecordObjectSync: retrieveRecordObjectSync,
            FormatDateEDM: formatDateEDM,
            ExecuteCustomActionSync: executeCustomActionSync,
            ExecuteCustomGlobalActionSync: executeCustomGlobalActionSync,
            ExecuteFetchXML: executeFetchXML
        };
    }());
}
