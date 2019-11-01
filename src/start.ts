import app from './app';
import { getConfig, exec } from './utils';

// Main wrapped for asyncness
const main = async () => {

    let afterStart = async () => {};

    if (process.argv.length < 4){
        console.log('Usage : <port> <config_file>')
        process.exit(2)
    }

    const port = Number(process.argv[2]);

    if (isNaN(port) || port > 65535) {
        throw new Error(`invalid port ${port}`)
    }

    const filename = process.argv[3];
    const minioInfos = await getConfig(filename);
    
    const command = process.argv[4];
    if (command){
        afterStart = async () => {
            await exec(command);
        }
    }
    
    await app(port, minioInfos, afterStart);
};

main();


