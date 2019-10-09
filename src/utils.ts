import { Stream } from 'stream';

export const asyncMiddleware = fn =>
  (req, res, next) => {
    Promise.resolve(fn(req, res, next))
      .catch(next);
  };


export const streamToBuffer = (stream: Stream) => {
    return new Promise<Buffer>((resolve, reject) => {
        let buffers = [];
        stream.on('error', reject);
        stream.on('data', (data) => buffers.push(data))
        stream.on('end', () => resolve(Buffer.concat(buffers)))
    });
}