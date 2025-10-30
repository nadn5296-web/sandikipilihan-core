(() => {
  console.log("✅ Dashboard script loaded");

  document.addEventListener("DOMContentLoaded", () => {
    const infoBox = document.querySelector("#dashboardInfo");
    if (!infoBox) return;

    infoBox.innerHTML = "<p>📊 Loading dashboard data...</p>";

    setTimeout(() => {
      infoBox.innerHTML = `
        <p><strong>Welcome to GH GROUP Dashboard!</strong></p>
        <p>Data updated at: ${new Date().toLocaleString()}</p>
      `;
    }, 500);
  });
})();
