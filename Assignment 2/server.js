const express = require('express');
const os = require('os');
const http = require('http');
const app = express();
const fs= require('fs');
const port = 3001;

app.set('views', __dirname);
app.set('view engine', 'ejs');

app.get('/info', (req, res) => {
    // Get device info
    const deviceInfo = {
      username: os.userInfo().username,
      homedir: os.userInfo().homedir,
      hostname: os.hostname(),
      release: os.release(),
      platform: os.platform(),
      type: os.type(),
      arch: os.arch(),
      cpus: os.cpus().length,
      uptime: os.uptime(),
    };

    fs.writeFile('info.txt', JSON.stringify(deviceInfo, null, 3 ), (err) => {
        if (err) throw err;
        console.log('The file saved!');
      }); 
  
      
    res.render('info', { deviceInfo });
});

/////////////////////////////
app.get('/download', (req, res) => {
  res.download('info.txt');
});

// Listen on port = 3001
app.listen(port, () => {
  console.log('Server is running ... ');
});