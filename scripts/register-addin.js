const fs = require('fs');
const path = require('path');
const os = require('os');

const manifestPath = path.resolve(__dirname, '../manifest.xml');
const manifestContent = fs.readFileSync(manifestPath, 'utf8');

// 获取用户目录
const homeDir = os.homedir();

// MacOS下的所有可能的Excel加载项注册目录
const possibleManifestLocations = [
    // 开发模式目录
    path.join(homeDir, 'Library/Containers/com.microsoft.Excel/Data/Documents/wef'),
    // 正式注册目录
    path.join(homeDir, 'Library/Group Containers/UBF8T346G9.Office/User Content.localized/Manifests.localized'),
    // 备用目录
    path.join(homeDir, 'Library/Group Containers/UBF8T346G9.Office/com.microsoft.Excel/com.microsoft.Excel.userManifests'),
    // Office 2021及更新版本
    path.join(homeDir, 'Library/Group Containers/UBF8T346G9.Office/com.microsoft.Office365V2/User Content/Manifests')
];

function ensureDirectoryExists(directory) {
    if (!fs.existsSync(directory)) {
        try {
            fs.mkdirSync(directory, { recursive: true });
            console.log(`Created directory: ${directory}`);
            return true;
        } catch (error) {
            console.error(`Failed to create directory ${directory}:`, error);
            return false;
        }
    }
    return true;
}

function registerManifest() {
    const manifestId = manifestContent.match(/<Id>(.*?)<\/Id>/)?.[1] || 'default';
    let registeredCount = 0;

    possibleManifestLocations.forEach(location => {
        try {
            console.log(`\nTrying to register in: ${location}`);
            
            if (!ensureDirectoryExists(location)) {
                console.log(`Skipping ${location} - cannot create directory`);
                return;
            }

            const targetPath = path.join(location, `${manifestId}.xml`);
            fs.copyFileSync(manifestPath, targetPath);
            console.log(`Successfully registered manifest at: ${targetPath}`);
            registeredCount++;
        } catch (error) {
            console.log(`Failed to register in ${location}:`, error.message);
        }
    });

    return registeredCount;
}

console.log('Starting Add-in Registration...\n');

const registeredCount = registerManifest();

if (registeredCount > 0) {
    console.log('\n✅ Add-in registration completed successfully!');
    console.log('\nNext steps:');
    console.log('1. Quit Excel completely if it\'s running');
    console.log('2. Start the dev server: npm start');
    console.log('3. Open Excel and check the Home tab for the add-in');
    console.log('\nIf the add-in is not visible:');
    console.log('1. Open Excel > Preferences');
    console.log('2. Go to Security & Privacy');
    console.log('3. Check "Trust access to the Office Add-ins platform"');
    console.log('4. Restart Excel');
} else {
    console.error('\n❌ Failed to register add-in in any location');
    console.log('\nTroubleshooting steps:');
    console.log('1. Ensure you have the necessary permissions');
    console.log('2. Try running with sudo if needed');
    console.log('3. Check if Excel is installed correctly');
    console.log('4. Verify the manifest file is valid');
}

// 输出注册位置供参考
console.log('\nManifest registration locations:');
possibleManifestLocations.forEach((location, index) => {
    console.log(`${index + 1}. ${location}`);
});
