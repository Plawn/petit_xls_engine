import app from './app';

// Main wrapped for asyncness
const main = () => {
    const port = Number(process.argv[2]);
    if (isNaN(port) || port > 65535) {
        throw new Error(`invalid port ${port}`)
    }
    app(port);
};

main();


