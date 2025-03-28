const companies = ["Apple", "Microsoft", "Google", "Amazon", "Facebook"];
const suggestionsDiv = document.getElementById("suggestions");
const companyInput = document.getElementById("companyInput");

companyInput.addEventListener("input", function() {
    const input = this.value.toLowerCase();
    suggestionsDiv.innerHTML = "";
    if (input.length >= 3) {
        const filteredCompanies = companies.filter(company =>
            company.toLowerCase().startsWith(input)
        );
        filteredCompanies.forEach(company => {
            const div = document.createElement("div");
            div.textContent = company;
            div.onclick = function() {
                companyInput.value = company;
                suggestionsDiv.innerHTML = "";
            };
            suggestionsDiv.appendChild(div);
        });
    }
});

function generateExcelFromTemplate() {
    const companyName = document.getElementById("companyInput").value;
    if (!companyName) {
        alert("Proszę wybrać nazwę firmy");
        return;
    }

    fetch("szablon.xlsx")
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: "array", cellStyles: true });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            // Zachowaj oryginalny styl komórki A5
            const originalStyle = worksheet["A5"].s;

            // Wprowadź nazwę firmy do komórki A5, zachowując styl
            worksheet["A5"] = { v: companyName, s: originalStyle };

            // Zapisz zmodyfikowany plik
            XLSX.writeFile(workbook, `Zlecenie_${companyName}.xlsx`, { cellStyles: true });
        })
        .catch(error => {
            console.error("Błąd podczas wczytywania szablonu:", error);
            alert("Wystąpił błąd podczas generowania pliku Excel.");
        });
}
