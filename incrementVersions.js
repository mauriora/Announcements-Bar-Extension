const fs = require('fs')
const packagePath = './package.json';
const solutionPath = './config/package-solution.json';

const versiony = require('versiony')
const chalk = require('chalk');

const updateFile = (file, version, nestedObject, arrayProperty) => {
    const newVersion = `${version}.0`;
    const targetName = `${chalk.cyan( file )}${nestedObject ? '.' + chalk.green( nestedObject ) : '' }`;
    console.log(`Updating ${targetName}.version=${chalk.yellow(newVersion)}`);
    const fileContent = require(file);
    const target = nestedObject ? fileContent[nestedObject] : fileContent;

    target.version = newVersion;
    if( typeof target[arrayProperty] === 'object' ) {
        for( const item of target[arrayProperty]) {
            console.log(`Updating ${targetName}.${chalk.blueBright( arrayProperty )}[${chalk.green( item['title'] ?? item['id'])}].version=${chalk.yellow(newVersion)}`);
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
    
    console.log(`Updated ${chalk.cyan( packagePath )} to ${chalk.yellow( newVersion.version )}`);
    updateFile( solutionPath, newVersion.version, 'solution', 'features' );
}

main();