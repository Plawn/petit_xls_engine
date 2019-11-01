import app from './app';
import { getConfig } from './utils';

// Main wrapped for asyncness
const main = async () => {

    if (process.argv.length != 4){
        console.log('Usage : <port> <config_file>')
        process.exit(2)
    }

    const port = Number(process.argv[2]);

    if (isNaN(port) || port > 65535) {
        throw new Error(`invalid port ${port}`)
    }

    const filename = process.argv[3];
    const minioInfos = await getConfig(filename);
    
    
    app(port, minioInfos, ()=> {
        process.stdout.write('started');
    });
};

main();


