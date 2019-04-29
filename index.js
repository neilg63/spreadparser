const express = require('express');
const app = express();
const api = require("./routes/api");

const PORT = 5555;

app.use("/api", api);


app.listen(PORT, () => {
  console.log(`Running on port ${PORT}`)
});