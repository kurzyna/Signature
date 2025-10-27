// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// === POLITYKA PODPISÓW (zmień wg potrzeby) =======================
const DEFAULT_TEMPLATE = "templateA"; // szablon firmowy
const NEW_MAIL_TEMPLATE = "templateA"; // nowa wiadomość
const REPLY_TEMPLATE = "templateA"; // odpowiedź: "templateA" albo "none"
const FORWARD_TEMPLATE = "templateA"; // przekazanie: "templateA" albo "none"
// ================================================================

function save_user_settings_to_roaming_settings() {
  Office.context.roamingSettings.saveAsync(function (asyncResult) {
    console.log("save_user_info_str_to_roaming_settings - " + JSON.stringify(asyncResult));
  });
}

function disable_client_signatures_if_necessary() {
  if ($("#checkbox_sig").prop("checked") === true) {
    Office.context.mailbox.item.disableClientSignatureAsync(function (asyncResult) {
      console.log("disable_client_signature_if_necessary - " + JSON.stringify(asyncResult));
    });
  }
}

function save_signature_settings() {
  const user_info_str = localStorage.getItem("user_info");
  if (!user_info_str) {
    console.warn("save_signature_settings: brak user_info w localStorage");
    return;
  }

  // Zapisz dane użytkownika + narzuć jeden szablon dla wszystkich typów wiadomości
  Office.context.roamingSettings.set("user_info", user_info_str);
  Office.context.roamingSettings.set("newMail", NEW_MAIL_TEMPLATE);
  Office.context.roamingSettings.set("reply", REPLY_TEMPLATE);
  Office.context.roamingSettings.set("forward", FORWARD_TEMPLATE);

  Office.context.roamingSettings.set("override_olk_signature", $("#checkbox_sig").prop("checked"));

  save_user_settings_to_roaming_settings();
  disable_client_signatures_if_necessary();

  $("#message").show("slow");
}

// ——— wstawianie treści ———
function set_body(str) {
  Office.context.mailbox.item.body.setAsync(
    get_cal_offset() + str,
    { coercionType: Office.CoercionType.Html },
    function (asyncResult) {
      console.log("set_body - " + JSON.stringify(asyncResult));
    }
  );
}

function set_signature(str) {
  Office.context.mailbox.item.body.setSignatureAsync(
    str,
    { coercionType: Office.CoercionType.Html },
    function (asyncResult) {
      console.log("set_signature - " + JSON.stringify(asyncResult));
    }
  );
}

function insert_signature(str) {
  if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Appointment) {
    set_body(str);
  } else {
    set_signature(str);
  }
}

// ——— jedyny przycisk testu ———
function test_template_A() {
  try {
    const ls = localStorage.getItem("user_info");
    if (ls) {
      _user_info = _user_info || JSON.parse(ls);
    }
  } catch (e) {
    console.warn("Nie można sparsować user_info:", e);
  }

  const str = get_template_A_str(_user_info || {});
  console.log("test_template_A -> length:", str?.length || 0);
  insert_signature(str);
}

// ——— nawigacja do edycji ———
function navigate_to_taskpane2() {
  window.location.href = "editsignature.html";
}
