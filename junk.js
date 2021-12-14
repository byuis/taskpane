console.log("loading junk----------------------------------")
console.log("starting--------------------------")




Office.onReady((info) => {
  console.log("at on ready")
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

