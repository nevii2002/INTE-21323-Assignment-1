async function debugResponse() {
    try {
        console.log("Fetching /api/staff/1/kpis...");
        const response = await fetch('http://localhost:5000/api/staff/1/kpis');

        console.log(`Status: ${response.status}`);

        if (response.ok) {
            const data = await response.json();
            console.log("Response Data:", JSON.stringify(data, null, 2));

            if (data.length > 0) {
                const first = data[0];
                if (first.EmployeeKPI) {
                    console.log("EmployeeKPI field found:", first.EmployeeKPI);
                    if (first.EmployeeKPI.customWeight !== undefined) {
                        console.log("customWeight field found:", first.EmployeeKPI.customWeight);
                    } else {
                        console.error("customWeight MISSING in EmployeeKPI!");
                    }
                } else {
                    console.error("EmployeeKPI field MISSING on KPI object!");
                }
            } else {
                console.log("No KPIs assigned to this user yet (Empty Array).");
            }
        } else {
            console.error("Request failed.");
        }
    } catch (e) {
        console.error("Fetch error:", e.message);
    }
}

debugResponse();
