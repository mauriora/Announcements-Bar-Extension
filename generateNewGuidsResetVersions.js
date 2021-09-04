const PACKAGE_VERSION = '0.0.1'
const SOLUTION_VERSION = '0.0.1.0'
const MANIFEST_VERSION = '*'

const crypto = require("crypto");
const fs = require('fs')
const path = require('path');
const packagePath = path.resolve(__dirname, './package.json');
const solutionPath = path.resolve(__dirname, './config/package-solution.json');
const extensionsPath = path.resolve(__dirname, './src/extensions');
const manifestEnding = '.manifest.json';


const updateFile = (file, id, version, nestedObject) => {
    console.log(`Updating ${path.basename(file)}${nestedObject ? '.' + nestedObject : '' } .id=${id} .version=${version} file=${file}`);
    const fileContent = require(file);

    if (undefined !== id) {
        if (nestedObject) {
            fileContent[nestedObject].id = id;
        } else {
            fileContent.id = id;
        }
    }
    if (undefined !== version) {
        if (nestedObject) {
            fileContent[nestedObject].version = version;
        } else {
            fileContent.version = version;
        }
    }
    fs.writeFileSync(file, JSON.stringify(fileContent, null, 4))
}


const getGuid = () => {
    const id = crypto.randomBytes(16).toString("hex");
    const guid = `${id.substr(0, 8)}-${id.substr(8, 4)}-${id.substr(12, 4)}-${id.substr(16, 4)}-${id.substr(20)}`

    return guid;
}

const setSolutionGuid = () => updateFile(solutionPath, getGuid(), SOLUTION_VERSION, 'solution' );

const ifManifestSetGuid = (folderName, fileName) =>
    fileName.endsWith(manifestEnding) &&
    updateFile(
        path.join(extensionsPath, folderName, fileName),
        getGuid(),
        MANIFEST_VERSION
    );

const processWebPartFolderFiles = (folderName, files) => files.forEach(file => ifManifestSetGuid(folderName, file));

const getWebPartFolderFiles = (folderName) =>
    fs.readdir(
        path.join(extensionsPath, folderName),
        (err, files) => processWebPartFolderFiles(folderName, files)
    );


const processWebPartsFolder = (err, webPartFolders) => webPartFolders.forEach(getWebPartFolderFiles);


const setManifestsGuids = () => fs.readdir(extensionsPath, processWebPartsFolder);

const setPackageVersion = () => updateFile(packagePath, undefined, PACKAGE_VERSION);


function main() {
    setPackageVersion();
    setSolutionGuid();
    setManifestsGuids();
}

main();