// Select DOM elements to work with
const signInButton = document.getElementById("btnSignIn");
const profileButton = document.getElementById("btnProfile");
const peopleButton = document.getElementById("btnPeople");
const groupsButton = document.getElementById("btnGroups");
const signOutButton = document.getElementById("btnSignOut");
const idTokenButton = document.getElementById("btnIDToken");
const tokenResponseButton = document.getElementById("btnTokenResponse");
const resultsButton = document.getElementById("btnResults");
const logButton = document.getElementById("btnLog");
const tabsDiv = document.getElementById("Tabs");
const idTokenTab = document.getElementById("idTokenTab");
const tokenResponseTab = document.getElementById("tokenResponseTab");
const resultsTab = document.getElementById("resultsTab");
const logTab = document.getElementById("logTab");

function setSignedIn() {
    signInButton.setAttribute("onclick", "");
    signInButton.setAttribute('class', "btn btn-secondary");

    profileButton.setAttribute("onclick", "getProfile()");
    profileButton.setAttribute('class', "btn btn-primary");

    peopleButton.setAttribute("onclick", "getPeople()");
    peopleButton.setAttribute('class', "btn btn-primary");

    groupsButton.setAttribute("onclick", "");
    groupsButton.setAttribute('class', "btn btn-primary");

    signOutButton.setAttribute("onclick", "signOut()");
    signOutButton.setAttribute('class', "btn btn-primary");

    idTokenButton.setAttribute("onclick", "switchTab('idTokenTab');");
    idTokenButton.setAttribute('class', "btn btn-primary");

    tokenResponseButton.setAttribute("onclick", "switchTab('tokenResponseTab');");
    tokenResponseButton.setAttribute('class', "btn btn-primary");

    resultsButton.setAttribute("onclick", "switchTab('resultsTab');");
    resultsButton.setAttribute('class', "btn btn-primary");

    logButton.setAttribute("onclick", "switchTab('logTab');");
    logButton.setAttribute('class', "btn btn-primary");
   
}

function setSignedOut() {
    signInButton.setAttribute("onclick", "signIn()");
    signInButton.setAttribute('class', "btn btn-primary");

    profileButton.setAttribute("onclick", "");
    profileButton.setAttribute('class', "btn btn-secondary");

    peopleButton.setAttribute("onclick", "");
    peopleButton.setAttribute('class', "btn btn-secondary");

    groupsButton.setAttribute("onclick", "");
    groupsButton.setAttribute('class', "btn btn-secondary");

    signOutButton.setAttribute("onclick", "");
    signOutButton.setAttribute('class', "btn btn-secondary");

    idTokenButton.setAttribute("onclick", "");
    idTokenButton.setAttribute('class', "btn btn-secondary");

    tokenResponseButton.setAttribute("onclick", "");
    tokenResponseButton.setAttribute('class', "btn btn-secondary");

    resultsButton.setAttribute("onclick", "");
    resultsButton.setAttribute('class', "btn btn-secondary");

    logButton.setAttribute("onclick", "");
    logButton.setAttribute('class', "btn btn-secondary");
   
}

function setAuthorityText() {
    const accountTypes = document.getElementById("accountTypes");
    const authorityText = document.getElementById("authorityText");

    if (accountTypes.selectedIndex === 0) {
        authorityText.innerHTML = "https://login.microsoftonline.com/common";
    }
    else if (accountTypes.selectedIndex === 1) {
        authorityText.innerHTML = "https://login.microsoftonline.com/organizations";
    }
    else if (accountTypes.selectedIndex === 2) {
        authorityText.innerHTML = "https://login.microsoftonline.com/consumers";
    }    
    else if (accountTypes.selectedIndex === 3) {
        authorityText.innerHTML = "https://login.microsoftonline.com/idfordevs.dev";
    }
    else {
        authorityText.innerHTML = accountTypes.selectedIndex;
    }
}

function updateIdTokenTab(claims, idToken) {

    idTokenTab.innerHTML = ''; 

    for (const key in claims) {
        var p = document.createElement('p');
        p.innerHTML = "<strong>" + key + ": </strong>" + claims[key];
        idTokenTab.appendChild(p);
    }

    if (idToken) {
        const jwtmsLink = "https://jwt.ms#id_token=" + idToken;
        const token = document.createElement('p');
        token.innerHTML = "<strong>ID Token: </strong>" + "<a target=\"_blank\" rel=\"noopener noreferrer\" href=\"https://jwt.ms#id_token=" + idToken + "\">idToken</a>" ;
        idTokenTab.appendChild(token);
    }  
    else {
        const fromCache = document.createElement('p');
        fromCache.innerHTML = "<strong>User auth from cache - MSAL had a valid ID token so \"fresh\" authentication not needed.</strong>";
        idTokenTab.appendChild(fromCache);
    }
}

function updateTokenResponseTab(response) {
    tokenResponseTab.innerHTML = ''; 

    for (const key in response) {
        if ( key == "account") {
            var p = document.createElement('p');
            p.innerHTML = "<strong>" + key + ": </strong>";
            tokenResponseTab.appendChild(p);
            const account = response.account;
            for (const akey in account ) {
                var p2 = document.createElement('p');
                p2.innerHTML = "<strong>" + akey + ": </strong>" + account[akey];
                p2.style.marginLeft = "25px";
                tokenResponseTab.appendChild(p2);
            }
        }
        else {
            var p = document.createElement('p');
            p.innerHTML = "<strong>" + key + ": </strong>" + response[key];    
            tokenResponseTab.appendChild(p);
        }
    }
}

function updateResultsTab(data) {
    resultsTab.innerHTML = ''; 

    for (const key in data) {
        var p = document.createElement('p');
        p.innerHTML = "<strong>" + key + ": </strong>" + data[key];
        resultsTab.appendChild(p);
    }
}


function updatePeopleResultsTab(data) {
    resultsTab.innerHTML = ''; 

    for (const key in data.value) {
        const person = data.value[key];
        for (const pkey in person ) {
            var p = document.createElement('p');
            p.innerHTML = "<strong>" + pkey + ": </strong>" + person[pkey];
            resultsTab.appendChild(p);
        }
    }
}

function switchTab(tabID) {
    var switchToTab = document.getElementById(tabID);
    for (var i = 0; i < tabsDiv.childNodes.length; i++) {
        var node = tabsDiv.childNodes[i];
        if (node.nodeType == 1) { 
            if (node == switchToTab){
                node.style.display = 'block';
            }
            else {
                node.style.display = 'none';
            }
        }
    }
}
