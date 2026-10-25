import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

const DISABLED_READABLE_STREAM_NAMESPACE = "read-only-zip-disabled-readable-stream";

const JSZIP_READ_ONLY_MODULES = new Map([
  ["flate.js", {
    sha256: "57f1452680cf7af6a9659de398d847dd1e870ea3a601662c498b073c640eb547",
    source: `"use strict";
var USE_TYPEDARRAY = (typeof Uint8Array !== "undefined") && (typeof Uint16Array !== "undefined") && (typeof Uint32Array !== "undefined");

var pako = require("pako");
var utils = require("./utils");
var GenericWorker = require("./stream/GenericWorker");

var ARRAY_TYPE = USE_TYPEDARRAY ? "uint8array" : "array";

exports.magic = "\\x08\\x00";

function FlateWorker(options) {
    GenericWorker.call(this, "FlateWorker/Inflate");

    this._pako = null;
    this._pakoOptions = options;
    this.meta = {};
}

utils.inherits(FlateWorker, GenericWorker);

FlateWorker.prototype.processChunk = function (chunk) {
    this.meta = chunk.meta;
    if (this._pako === null) {
        this._createPako();
    }
    this._pako.push(utils.transformTo(ARRAY_TYPE, chunk.data), false);
};

FlateWorker.prototype.flush = function () {
    GenericWorker.prototype.flush.call(this);
    if (this._pako === null) {
        this._createPako();
    }
    this._pako.push([], true);
};

FlateWorker.prototype.cleanUp = function () {
    GenericWorker.prototype.cleanUp.call(this);
    this._pako = null;
};

FlateWorker.prototype._createPako = function () {
    this._pako = new pako.Inflate({
        raw: true,
        level: this._pakoOptions.level || -1
    });
    var self = this;
    this._pako.onData = function(data) {
        self.push({
            data : data,
            meta : self.meta
        });
    };
};

exports.uncompressWorker = function () {
    return new FlateWorker({});
};
`,
  }],
  ["compressions.js", {
    sha256: "50e650a7eaf2a54ede299c12dfba30a362c7644e7c00df0d16ab3cf228bb77b0",
    source: `"use strict";

var GenericWorker = require("./stream/GenericWorker");

exports.STORE = {
    magic: "\\x00\\x00",
    uncompressWorker : function () {
        return new GenericWorker("STORE decompression");
    }
};
exports.DEFLATE = require("./flate");
`,
  }],
  ["compressedObject.js", {
    sha256: "b91f823d32f39cda4002bc0954993f20ffea6747d42109253b7d6347b4abb31b",
    source: `"use strict";

var external = require("./external");
var DataWorker = require("./stream/DataWorker");
var DataLengthProbe = require("./stream/DataLengthProbe");

function CompressedObject(compressedSize, uncompressedSize, crc32, compression, data) {
    this.compressedSize = compressedSize;
    this.uncompressedSize = uncompressedSize;
    this.crc32 = crc32;
    this.compression = compression;
    this.compressedContent = data;
}

CompressedObject.prototype = {
    getContentWorker: function () {
        var worker = new DataWorker(external.Promise.resolve(this.compressedContent))
            .pipe(this.compression.uncompressWorker())
            .pipe(new DataLengthProbe("data_length"));

        var that = this;
        worker.on("end", function () {
            if (this.streamInfo["data_length"] !== that.uncompressedSize) {
                throw new Error("Bug : uncompressed data size mismatch");
            }
        });
        return worker;
    }
};

module.exports = CompressedObject;
`,
  }],
]);

const ZIP_OBJECT_WRITER_BLOCK = /    \/\*\*\r?\n     \* Return a worker for the compressed content\.\r?\n     \* @private\r?\n     \* @param \{Object\} compression the compression object to use\.\r?\n     \* @param \{Object\} compressionOptions the options to use when compressing\.\r?\n     \* @return Worker the worker\.\r?\n     \*\/\r?\n    _compressWorker: function \(compression, compressionOptions\) \{[\s\S]*?\r?\n    \},\r?\n/;

const ZIP_OBJECT_SHA256 = "18871af92fac75d015ba7d4d3105c0746839ab089f78670cb86d76bfa53b1ce6";

function sha256(source) {
  return createHash("sha256").update(source).digest("hex");
}

async function loadPrunedModule(moduleName, filePath) {
  const original = await fs.readFile(filePath, "utf8");
  const actual = sha256(original);
  if (moduleName === "zipObject.js") {
    if (actual !== ZIP_OBJECT_SHA256) {
      throw new Error(`Pinned JSZip source changed: ${moduleName} (${actual}).`);
    }
    if (!ZIP_OBJECT_WRITER_BLOCK.test(original)) {
      throw new Error(`Pinned JSZip writer block not found in ${moduleName}.`);
    }
    return original.replace(ZIP_OBJECT_WRITER_BLOCK, "");
  }
  const replacement = JSZIP_READ_ONLY_MODULES.get(moduleName);
  if (!replacement) {
    throw new Error(`No read-only replacement registered for ${moduleName}.`);
  }
  if (actual !== replacement.sha256) {
    throw new Error(`Pinned JSZip source changed: ${moduleName} (${actual}).`);
  }
  return replacement.source;
}

export function createReadOnlyZipPlugin(root, options = {}) {
  return {
    name: "read-only-zip",
    setup(build) {
      build.onResolve({ filter: /^jszip$/ }, () => ({
        path: path.resolve(root, "src/ooxml/jszipReader.ts"),
      }));
      build.onResolve({ filter: /^pako$/ }, () => ({
        path: path.resolve(root, "node_modules/pako/lib/inflate.js"),
      }));
      build.onResolve({ filter: /^readable-stream$/ }, (args) => {
        const importer = args.importer.replaceAll("\\", "/");
        if (!/(?:^|\/)node_modules\/jszip\/lib\/(?:support|nodejs\/NodejsStreamOutputAdapter)\.js$/.test(importer)) {
          return undefined;
        }
        return {
          path: "jszip-readable-stream-disabled",
          namespace: DISABLED_READABLE_STREAM_NAMESPACE,
        };
      });
      build.onLoad({ filter: /.*/, namespace: DISABLED_READABLE_STREAM_NAMESPACE }, (args) => {
        const contents = "module.exports = Object.freeze({});";
        options.onVirtualLoad?.({
          namespace: args.namespace,
          path: args.path,
          suffix: args.suffix,
          importAttributes: args.with,
          loader: "js",
          resolveDir: "",
          contents,
        });
        return { contents, loader: "js" };
      });
      build.onLoad({ filter: /jszip[\\/]lib[\\/](flate|compressions|compressedObject|zipObject)\.js$/ }, async (args) => {
        const moduleName = path.basename(args.path);
        const contents = await loadPrunedModule(moduleName, args.path);
        return {
          contents,
          loader: "js",
          resolveDir: path.dirname(args.path),
        };
      });
    },
  };
}
