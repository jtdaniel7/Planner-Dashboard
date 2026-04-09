(async () => {
  const res = await fetch('https://httpbin.org/get');
  console.log('Status:', res.status);
  const data = await res.json();
  console.log('Origin IP:', data.origin);
})();
