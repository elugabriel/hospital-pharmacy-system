document.addEventListener("DOMContentLoaded", function () {

    let allDrugs = [];

    const searchInput = document.getElementById("drug-search");
    const suggestionBox = document.getElementById("drug-suggestions");
    const confirmBtn = document.getElementById("confirm-payment");
    const cancelBtn = document.getElementById("cancel-sale");
    const cartTable = document.querySelector("#cart-table tbody");

    // -------------------- FETCH DRUGS --------------------
    async function fetchDrugs(query = "") {
        try {
            const url = query ? `/api/drugs?q=${query}` : "/api/drugs";
            const res = await fetch(url);
            const data = await res.json();
            allDrugs = data;
            return data;
        } catch (err) {
            console.error("Fetch drugs error:", err);
            return [];
        }
    }

    // Initial fetch
    fetchDrugs();

    // -------------------- DRUG SEARCH --------------------
    searchInput.addEventListener("input", async function () {
        const val = this.value.trim().toLowerCase();
        suggestionBox.innerHTML = "";

        if (!val) {
            suggestionBox.classList.add("d-none");
            return;
        }

        const matches = await fetchDrugs(val);
        const filtered = matches.filter(d => d.name.toLowerCase().startsWith(val));

        filtered.forEach(d => {
            const li = document.createElement("li");
            li.className = "list-group-item list-group-item-action";
            li.textContent = `${d.name} (${d.strength}) - ₦${d.unit_price}`;
            li.style.cursor = "pointer";

            li.addEventListener("click", function () {
                addDrugToCart(d);
                searchInput.value = "";
                suggestionBox.classList.add("d-none");
            });

            suggestionBox.appendChild(li);
        });

        suggestionBox.classList.toggle("d-none", filtered.length === 0);
    });

    // Click outside to close dropdown
    document.addEventListener("click", function (e) {
        if (!searchInput.contains(e.target) && !suggestionBox.contains(e.target)) {
            suggestionBox.classList.add("d-none");
        }
    });

    // -------------------- ADD DRUG TO CART --------------------
    function addDrugToCart(drug) {
        // Check if already in cart
        let existingRow = Array.from(cartTable.rows).find(row => row.dataset.id == drug.id);
        if (existingRow) {
            const qtyCell = existingRow.querySelector(".drug-qty");
            qtyCell.textContent = parseInt(qtyCell.textContent) + 1;
        } else {
            const row = document.createElement("tr");
            row.dataset.id = drug.id;
            row.innerHTML = `
                <td>${drug.name}</td>
                <td>${drug.strength}</td>
                <td>${parseFloat(drug.unit_price).toFixed(2)}</td>
                <td class="drug-qty">1</td>
                <td class="drug-subtotal">${parseFloat(drug.unit_price).toFixed(2)}</td>
                <td><button class="btn btn-sm btn-danger remove-btn">Remove</button></td>
            `;
            cartTable.appendChild(row);

            // Remove button
            row.querySelector(".remove-btn").addEventListener("click", function () {
                row.remove();
                updateTotals();
            });
        }
        updateTotals();
    }

    // -------------------- UPDATE TOTALS --------------------
    function updateTotals() {
        let subtotal = 0;

        Array.from(cartTable.rows).forEach(row => {
            const qty = parseInt(row.querySelector(".drug-qty").textContent);
            const unitPrice = parseFloat(row.cells[2].textContent);
            const lineTotal = qty * unitPrice;
            row.querySelector(".drug-subtotal").textContent = lineTotal.toFixed(2);
            subtotal += lineTotal;
        });

        const discount = parseFloat(document.getElementById("discount").value) || 0;
        const tax = parseFloat(document.getElementById("tax").value) || 0;
        const grandTotal = subtotal - discount + tax;

        document.getElementById("subtotal").textContent = subtotal.toFixed(2);
        document.getElementById("grand-total").textContent = grandTotal.toFixed(2);
    }

    document.getElementById("discount").addEventListener("input", updateTotals);
    document.getElementById("tax").addEventListener("input", updateTotals);

    // -------------------- CONFIRM PAYMENT --------------------
    confirmBtn.addEventListener("click", async function () {

        if (cartTable.rows.length === 0) {
            alert("Cart is empty.");
            return;
        }
    
        const patient = prompt("Enter Patient Name or ID:");
        if (!patient) return;
    
        const items = Array.from(cartTable.rows).map(row => ({
            drug_name: row.cells[0].textContent,
            strength: row.cells[1].textContent,
            unit_price: parseFloat(row.cells[2].textContent),
            quantity: parseInt(row.cells[3].textContent)
        }));
    
        const payload = {
            patient_name: patient,
            patient_id: patient,
            items: items,
            subtotal: parseFloat(document.getElementById("subtotal").textContent),
            discount: parseFloat(document.getElementById("discount").value) || 0,
            tax: parseFloat(document.getElementById("tax").value) || 0,
            grand_total: parseFloat(document.getElementById("grand-total").textContent)
        };
    
        try {
            const res = await fetch("/pharmacy/confirm-payment", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)
            });
    
            if (!res.ok) throw new Error("Server error");
    
            const result = await res.json();
    
            if (result.success) {
                window.location.href = `/pharmacy/receipt/${result.receipt_id}`;
            } else {
                alert("Failed to save receipt.");
            }
    
        } catch (err) {
            console.error(err);
            alert("Error connecting to server.");
        }
    });
    

    // -------------------- CANCEL SALE --------------------
    cancelBtn.addEventListener("click", function () {
        if (confirm("Are you sure you want to cancel this sale?")) {
            cartTable.innerHTML = "";
            document.getElementById("subtotal").textContent = "0.00";
            document.getElementById("grand-total").textContent = "0.00";
            document.getElementById("discount").value = 0;
            document.getElementById("tax").value = 0;
        }
    });

});
