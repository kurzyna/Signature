// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

let _display_name;
let _job_title;
let _phone_number;
let _email_id;
let _department_name;
let _team_name;
let _message;

Office.initialize = function (reason) {
  on_initialization_complete();
};

function on_initialization_complete() {
  $(document).ready(function () {
    _output = $("textarea#output");
    _display_name = $("input#display_name");
    _email_id = $("input#email_id");
    _job_title = $("input#job_title");
    _phone_number = $("input#phone_number");
    _department_name = $("input#department_name");
    _team_name = $("input#team_name");
    _message = $("p#message");

    prepopulate_from_userprofile();
    load_saved_user_info();
  });
}

function prepopulate_from_userprofile() {
  const profile = Office.context.mailbox.userProfile || {};
  const displayName = (profile.displayName || "").trim().replace(/\s+/g, " ");
  const parts = displayName.split(" ").filter(Boolean);

  let finalName = displayName;

  if (parts.length >= 2) {
    const firstName = parts[parts.length - 1];
    const lastName = parts.slice(0, -1).join(" ");
    finalName = `${firstName} ${lastName}`;
  }

  _display_name.val(finalName);
  _email_id.val(profile.emailAddress || "");
}

function load_saved_user_info() {
  let user_info_str = localStorage.getItem("user_info");
  if (!user_info_str) {
    user_info_str = Office.context.roamingSettings.get("user_info");
  }

  if (user_info_str) {
    const user_info = JSON.parse(user_info_str);

    _display_name.val(user_info.name);
    _email_id.val(user_info.email);
    _job_title.val(user_info.job);
    _phone_number.val(user_info.phone);
    _department_name.val(user_info.department);
    _team_name.val(user_info.team);
  }
}

function display_message(msg) {
  _message.text(msg);
}

function clear_message() {
  _message.text("");
}

function is_not_valid_text(text) {
  return text.length <= 0;
}

function is_not_valid_email_address(email_address) {
  let email_address_regex = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/;
  return is_not_valid_text(email_address) || !email_address_regex.test(email_address);
}

function form_has_valid_data(name, email) {
  if (is_not_valid_text(name)) {
    display_message("Proszę podać prawidłowe imię i nazwisko.");
    return false;
  }

  if (is_not_valid_email_address(email)) {
    display_message("Proszę podać prawidłowy adres email.");
    return false;
  }

  return true;
}

function navigate_to_taskpane_assignsignature() {
  window.location.href = "assignsignature.html";
}

// === Inline walidacja ===
function showFieldError($input, msg) {
  clearFieldError($input);
  $input.addClass("invalid");
  $input.after('<div class="error-msg">' + msg + "</div>");
}
function clearFieldError($input) {
  $input.removeClass("invalid");
  const $n = $input.next(".error-msg");
  if ($n.length) $n.remove();
}

function validate_form_fields() {
  let ok = true;

  const $name = _display_name; // $("#display_name")
  const $email = _email_id; // $("#email_id")
  const $job = _job_title; // $("#job_title")  (u Ciebie jest required w HTML)

  // wyczyść stare błędy
  [$name, $email, $job].forEach(clearFieldError);

  const name = ($name.val() || "").trim();
  const email = ($email.val() || "").trim();
  const job = ($job.val() || "").trim();

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/i;

  if (!name) {
    showFieldError($name, "Podaj imię i nazwisko.");
    ok = false;
  }
  if (!email || !emailRegex.test(email)) {
    showFieldError($email, "Podaj poprawny adres e-mail.");
    ok = false;
  }
  // pole "Stanowisko" traktujemy jako wymagane (masz required w HTML)
  if (!job) {
    showFieldError($job, "Podaj stanowisko.");
    ok = false;
  }

  if (!ok) {
    $("#message").text("Uzupełnij brakujące pola.").show();
  } else {
    $("#message").hide();
  }
  return ok;
}

// czyść błędy w trakcie wpisywania
$(document).on("input", "#display_name, #email_id, #job_title", function () {
  clearFieldError($(this));
  $("#message").hide();
});

// --- zaktualizuj create_user_info ---
function create_user_info() {
  clear_message();
  if (!validate_form_fields()) return;

  let user_info = {
    name: _display_name.val().trim(),
    email: _email_id.val().trim(),
    job: _job_title.val().trim(),
    phone: _phone_number.val().trim(),
    department: _department_name.val().trim(),
    team: _team_name.val().trim(),
  };

  localStorage.setItem("user_info", JSON.stringify(user_info));
  navigate_to_taskpane_assignsignature();
}

function clear_all_fields() {
  _display_name.val("");
  _email_id.val("");
  _phone_number.val("");
  _job_title.val("");
  _department_name.val("");
  _team_name.val("");
}

function clear_all_localstorage_data() {
  localStorage.removeItem("user_info");
  localStorage.removeItem("newMail");
  localStorage.removeItem("reply");
  localStorage.removeItem("forward");
  localStorage.removeItem("override_olk_signature");
}

function clear_roaming_settings() {
  Office.context.roamingSettings.remove("user_info");
  Office.context.roamingSettings.remove("newMail");
  Office.context.roamingSettings.remove("reply");
  Office.context.roamingSettings.remove("forward");
  Office.context.roamingSettings.remove("override_olk_signature");

  Office.context.roamingSettings.saveAsync(function (asyncResult) {
    console.log("clear_roaming_settings - " + JSON.stringify(asyncResult));

    let message =
      "Wszystkie ustawienia zostały pomyślnie zresetowane! Ten dodatek nie będzie już wstawiać żadnych podpisów. Możesz teraz zamknąć to okno.";

    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      message = "Reset nie powiódł się. Spróbuj ponownie.";
    }

    display_message(message);
  });
}

function reset_all_configuration() {
  clear_all_fields();
  clear_all_localstorage_data();
  clear_roaming_settings();
}
