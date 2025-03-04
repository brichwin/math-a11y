function getHeadingForTable(table) {
    let parent = table.parentElement;
    
    let currentElement = table;
    let heading = null;
    
    while (currentElement) {
      currentElement = currentElement.previousElementSibling;
      
      if (currentElement) {
        if (currentElement.classList.contains('markdown-heading') || 
            currentElement.querySelector('h1, h2, h3, h4, h5, h6')) {
          heading = currentElement;
          break;
        }
      } else {
        if (parent.parentElement) {
          currentElement = parent;
          parent = parent.parentElement;
        } else {
          break; 
        }
      }
    }
    
    if (heading) {
      let headingElement = heading.querySelector('h1, h2, h3, h4, h5, h6');
      return headingElement ? headingElement.textContent : null;
    }
    
    return null;
  }
  

function generateAutoCorrectEntries() {
    // Get all tables on the page
    const tables = document.querySelectorAll('table');
    const entries = [];
    
    // Process each table
    tables.forEach(table => {
        // Skip the first row (the column headers)
        const rows = Array.from(table.querySelectorAll('tr')).slice(1);
        const categoryHeading = getHeadingForTable(table) || 'Missing'

        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            if (cells.length >= 6) { // Ensure we have enough cells
                // Extract data from cells
                const name = cells[0].textContent.trim();
                const symbol = cells[1].textContent.trim();
                const existingEntry = cells[3].textContent.trim();
                const proposedEntries = cells[4].textContent.trim();
                
                // Skip header or divider rows
                if (name === "-" || name === "----" || !name) return;
                
                // Process existing entries
                let existingCode = null;
                if (existingEntry !== "None" && existingEntry !== "N/A" && existingEntry !== "N\\A") {
                    // Extract the first entry code
                    const match = existingEntry.match(/\\[a-zA-Z0-9]+/);
                    if (match) {
                        existingCode = match[0];
                    }
                }
                
                // Process proposed entries
                const proposedCodes = proposedEntries.match(/\\[a-zA-Z0-9]+/g) || [];
                
                // Create an entry for each proposed code
                proposedCodes.forEach(code => {
                    entries.push({
                        category: categoryHeading,
                        name: code.substring(1), // Remove the leading backslash
                        symbol: symbol,
                        existingEntry: existingCode ? `"\\${existingCode}"` : '""'
                    });
                });
            }
        });
    });
    
    // Generate C# code
    let csharpCode = '// Extracted Auto-Correct Entries\n';
    let currentCategory = '';
    
    entries.forEach((entry, index) => {
        // Add a comment for the first entry of each symbol to group them
        if (entry.category !== currentCategory) {
            csharpCode += `\n// ${entry.category}\n`;
            currentCategory = entry.category
        }
        csharpCode += `new AutoCorrectEntry { Name = "${entry.name}", Symbol = "${entry.symbol}", ExistingEntry = ${entry.existingEntry} },\n`;
    });
    
    // Create a temporary textarea to allow copying to clipboard
    const textarea = document.createElement('textarea');
    textarea.value = csharpCode;
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand('copy');
    document.body.removeChild(textarea);
    
    console.log("C# code copied to clipboard");
}

