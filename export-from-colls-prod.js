/*
To download the table from colls-prod, while logged in.

Press F12 (or right-click anywhere on the page and select Inspect), then select the Console tab.

Paste the snippet below into the console prompt and press Enter:
*/

(async function scrapeAllPages() {
  let masterData = [];
  let pageCount = 1;

  function grabCurrentRows() {
    let rows = document.querySelectorAll("table tr");
    rows.forEach((row, index) => {
      // Include the table header row only on Page 1
      if (pageCount > 1 && index === 0) return;
      let cells = Array.from(row.querySelectorAll("th, td")).map(cell => {
        let text = cell.innerText.replace(/\n+/g, " ").replace(/"/g, '""').trim();
        return `"${text}"`;
      });
      if (cells.length > 0) masterData.push(cells.join(","));
    });
  }

  while (true) {
    console.log(`Scraping page ${pageCount}...`);
    grabCurrentRows();

    // Target the exact Next button ID found in DevTools
    let nextBtn = document.getElementById("course-results-next-page");
    
    // Stop when button is missing, disabled, or unreachable
    if (!nextBtn || nextBtn.disabled || nextBtn.getAttribute('aria-disabled') === 'true') {
      console.log(`Reached the end at page ${pageCount}.`);
      break;
    }

    nextBtn.click();
    pageCount++;
    
    // Wait 2.5 seconds for AJAX table reload before scraping the next set
    await new Promise(resolve => setTimeout(resolve, 2500));
  }

  // Generate and trigger single CSV download
  let csvContent = masterData.join("\n");
  let blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  let url = URL.createObjectURL(blob);
  let link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute("download", `Hamilton_Courses_All_Pages_${new Date().toISOString().slice(0,10)}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  console.log(`Done! Downloaded ${masterData.length - 1} total courses.`);
})();