const API_KEY = 'c2bf75e1-b226-4c99-9689-9139f9c8305e';
const auth = Buffer.from('key:' + API_KEY).toString('base64');
(async () => {
  const res = await fetch('https://api.redtailtechnology.com/crm/v1/contacts?keyword=Benedict', {
    headers: { 'Authorization': 'Basic ' + auth, 'Accept': 'application/json' }
  });
  console.log('Status:', res.status);
  const data = await res.json();
  console.log(JSON.stringify(data, null, 2).slice(0, 2000));
})();
