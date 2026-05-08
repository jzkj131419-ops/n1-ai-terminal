const express = require('express');
const app = express();
const port = 3000;

app.get('/', (req, res) => {
  res.send('AI Gateway is running! 🚀');
});

app.listen(port, () => {
  console.log(`AI Gateway listening on port ${port}`);
});
