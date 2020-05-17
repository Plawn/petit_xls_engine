import { Stream } from 'stream';
import { Response, Request, NextFunction } from 'express';

type ExpressHandler = (req: Request, res: Response, next: NextFunction) => Promise<any>

export const asyncMiddleware = (fn: ExpressHandler) =>
  (req: Request, res: Response, next: NextFunction) => {
    Promise
      .resolve(fn(req, res, next))
      .catch(next);
  };

export const portFromUrl = (host: string) => {
  let i = host.indexOf(':');
  return Number(host.slice(i + 1).split('/')[0]);
}


export const streamToBuffer = (stream: Stream) => {
  return new Promise<Buffer>((resolve, reject) => {
    let buffers = [];
    stream.on('error', reject);
    stream.on('data', (data) => buffers.push(data))
    stream.on('end', () => resolve(Buffer.concat(buffers)))
  });
}

export class SafeMap<T, U> extends Map<T, U> {
  _get: (<T>(key: T) => U) = Map.prototype.get
  get = (key: T) => {
    const res = this._get(key);
    if (!res) throw new Error(`${key} not found`);
    return res;
  }
}
