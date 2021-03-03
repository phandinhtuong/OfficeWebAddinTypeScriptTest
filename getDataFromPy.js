const {spawn} = require('child_process');
// const python = spawn('python', ['../../scrapyProj/crawlTyGia/crawlTyGia/spiders/sendDataToNode.py']);
const python = spawn('python', ['sendDataToNode.py'], {detached: true});
python.stdout.on('data', (data) => {
    // Do something with the data returned from python script
    console.log(data.toString());
});