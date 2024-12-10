microsoftTeams.appInitialization.notifySuccess();
const loading = document.getElementById("loading");
const successMsg = document.getElementById("success-msg");
const errorMsg = document.getElementById("error-msg");
loading.style.display = "block";

microsoftTeams.app.initialize().then(function () {
  microsoftTeams.app.getContext().then(function (context) {
    if (!context) {
      errorMsg.style.display = "block";
      return;
    }
    const userEmail = context.user?.userPrincipalName;

    fetch("https://api.test.cbrpgrd.com/external/api/teams_email_verify", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer test_token",
      },
      body: JSON.stringify({ email: userEmail }),
    })
      .then((response) => {
        loading.style.display = "none";
        if (response.ok) {
          successMsg.style.display = "block";
        } else {
          errorMsg.style.display = "block";
        }
      })
      .catch((error) => {
        console.error("Error verifying user:", error);
        errorMsg.style.display = "block";
      });
  });
});
