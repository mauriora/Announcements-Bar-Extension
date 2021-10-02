const fs = require('fs')
const packagePath = './package.json';
const solutionPath = './config/package-solution.json';

const versiony = require('versiony')


const updateFile = (file, version, nestedObject, arrayProperty) => {
    const newVersion = `${version}.0`;
    const targetName = `${file}${nestedObject ? '.' + nestedObject : '' }`;
    console.log(`Updating ${targetName}.version=${newVersion}`);
    const fileContent = require(file);
    const target = nestedObject ? fileContent[nestedObject] : fileContent;

    target.version = newVersion;
    if( typeof target[arrayProperty] === 'object' ) {
        console.log(`Updating ${target[arrayProperty].length} in ${targetName}.${arrayProperty} `);

        for( const item of target[arrayProperty]) {
            console.log(`Updating ${targetName}.${arrayProperty}[${item['title'] ?? item['id']}]`);
            item['version'] = newVersion;
        }
    } else {
        console.error(`${targetName}.${arrayProperty} is of type ${typeof target[arrayProperty]}`)
    }

    fs.writeFileSync(file, JSON.stringify(fileContent, null, 4))
}

function main() {
    const newVersion = versiony
        .patch()
        .with(packagePath)
        .end({quiet: true});
    
    console.log(`Updated ${packagePath} to ${newVersion.version}`);
    updateFile( solutionPath, newVersion.version, 'solution', 'features' );
}

main();