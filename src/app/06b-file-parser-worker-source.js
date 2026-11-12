import __fileParserWorkerSource from "../workers/file-parser.worker.js?worker-source";

/* The file parser worker's bundled source, started as a Blob worker. Kept in
   its own module so the worker itself can import the modules below this one. */
const fileParserWorkerSource = __fileParserWorkerSource;

export { fileParserWorkerSource };
