const fs = require('fs');
const JSZip = require('jszip');

console.log("foo");

async function getChromeDriverVersion(channel)
{
    // first, get version we want for the channel
    const channelResponse = await fetch("https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions.json");

    if (!channelResponse.ok)
    {
        console.log(`failed to fatch channel response: ${channelResponse}`);
        return;
    }

    const channelJson = await channelResponse.json();

    const version = channelJson.channels[channel].version;

    const response = await fetch("https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"); 

    if (!response.ok)
    {
        console.log(`failed to fetch: ${response}`);
        return;
    }

    const json = await response.json();

    const versionInfo = json.versions.filter((info)=>info.version.startsWith(version));

    if (versionInfo.length == 0)
    {
        console.log(`could not find version info for ${version}`);
        return;
    }

    console.log(`chromedriver: ${versionInfo[0]}`);

    const chromedriver = versionInfo[0].downloads.chromedriver.filter((download)=>download.platform == "win64");

    if (chromedriver.length == 0)
    {
        console.log("could not find chromeDriver entry for win64");
        return;
    }

    const url = chromedriver[0].url;

    console.log(`download info: ${versionInfo[0].version}: ${chromedriver[0].url}`);

    // now download the file
    const downloadResponse = await fetch(url);

    if (!downloadResponse.ok)
    {
        console.log(`url ${url} download failed: ${downloadResponse}`);
        return;
    }

    const buffer = await downloadResponse.blob();

    const filename = `${process.env.TEMP}\\chromedriver-stable.zip`;

    fs.writeFileSync(filename, await buffer.bytes());

    const zip = await JSZip.loadAsync(await buffer.bytes());

    if (!zip.files["chromedriver-win64/chromedriver.exe"])
    {
        console.log("can't find chromedriver.exe in zip package")
        return;
    }

    const fileData = await zip.files["chromedriver-win64/chromedriver.exe"].async('nodebuffer');

    fs.writeFileSync("..\\chromedriver.exe", fileData);
}

// 
// // .then((response)=>console.log(response.body.getReader().read().then((content)=>console.log(content))))

getChromeDriverVersion("Stable");
