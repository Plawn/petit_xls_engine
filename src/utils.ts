import { Stream } from 'stream';
import { promisify } from 'util';
import fs from 'fs';
import { minioInfosType } from './types';
import YAML from 'js-yaml';

const readFileAsync = promisify(fs.readFile);

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

export const getConfig = async (filename: string) => {
  const content = await readFileAsync(filename, 'utf-8');
  const parsed = YAML.load(content);
  const minioInfos: minioInfosType = {
    endpoint: parsed['MINIO_HOST'],
    passkey: parsed['MINIO_PASS'],
    access_key: parsed['MINIO_KEY'],
  }
  return minioInfos;
}
