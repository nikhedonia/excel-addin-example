const app = require('express')();

app.use(require('cors')());
app.get('/', (req,res)=> {
  res.json({cell: 'F6', value: Math.random()});
})


app.listen(3001);


