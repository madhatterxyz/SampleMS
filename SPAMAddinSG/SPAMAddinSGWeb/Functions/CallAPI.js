// The following string is an email address come from API Config
function getEmailAddress() {
    var xhr_object = new XMLHttpRequest();
    xhr_object.open("GET", "../api/Config/email", false);
    xhr_object.send();
    if (xhr_object.readyState == 4) {
        return JSON.parse(xhr_object.response);
    } else {
        return
        {
            Success: "Error",
            Data: "xhr_object.readyState is " + xhr_object.readyState,
        };
    }
}

// The following string is a body email come from API Config
function getBodyEmail() {
    var xhr_object = new XMLHttpRequest();
    xhr_object.open("GET", "../api/Config/bodyemail", false);
    xhr_object.send();
    if (xhr_object.readyState == 4) {
        return JSON.parse(xhr_object.response);
    } else {
        return
        {
            Success: "Error",
            Data: "xhr_object.readyState is " + xhr_object.readyState,
        };
    }
}