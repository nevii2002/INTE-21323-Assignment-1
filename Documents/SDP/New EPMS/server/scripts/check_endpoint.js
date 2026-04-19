async function checkEndpoint() {
    try {
        const response = await fetch('http://localhost:5000/api/staff/1/kpis');
        console.log(`Status Code: ${response.status}`);

        if (response.status === 404) {
            console.log('Endpoint NOT Found. Server might need restart.');
        } else if (response.status === 200) {
            console.log('Endpoint Found.');
        } else {
            console.log('Endpoint exists but returned ' + response.status);
        }
    } catch (error) {
        console.error('Connection Error:', error.message);
    }
}

checkEndpoint();
