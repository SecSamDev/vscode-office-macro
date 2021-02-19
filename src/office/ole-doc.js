const fs = require('fs').promises

const OleHeaderID = Buffer.from('D0CF11E0A1B11AE1', 'hex');


class OleHeader {
    constructor() {

    }

    /**
     * 
     * @param {Buffer} buffer 
     */
    load(buffer) {
        let i;
        for (i = 0; i < 8; i++) {
            if (OleHeaderID[i] != buffer[i])
                throw new Error('Not a valid compound document.')
        }

        this.secSize = 1 << buffer.readInt16LE(30);  // Size of sectors
        this.shortSecSize = 1 << buffer.readInt16LE(32);  // Size of short sectors
        this.SATSize = buffer.readInt32LE(44);  // Number of sectors used for the Sector Allocation Table
        this.dirSecId = buffer.readInt32LE(48);  // Starting Sec ID of the directory stream
        this.shortStreamMax = buffer.readInt32LE(56);  // Maximum size of a short stream
        this.SSATSecId = buffer.readInt32LE(60);  // Starting Sec ID of the Short Sector Allocation Table
        this.SSATSize = buffer.readInt32LE(64);  // Number of sectors used for the Short Sector Allocation Table
        this.MSATSecId = buffer.readInt32LE(68);  // Starting Sec ID of the Master Sector Allocation Table
        this.MSATSize = buffer.readInt32LE(72);  // Number of sectors used for the Master Sector Allocation Table

        // The first 109 sectors of the MSAT
        this.partialMSAT = new Array(109);
        for (i = 0; i < 109; i++)
            this.partialMSAT[i] = buffer.readInt32LE(76 + i * 4);
    }
}

const SEC_ID_FREE = -1;
const SEC_ID_END_OF_CHAIN = -2;
const SEC_ID_SAT = -3;
const SEC_ID_MSAT = -4;

class AllocationTable {
    /**
     * 
     * @param {OleCompoundDoc} doc 
     */
    constructor(doc) {
        this.doc = doc
    }
    async load(secIds) {
        this.table = new Array(secIds.length * (this.doc.header.secSize / 4))
        let buffer = await this.doc.readSectors(secIds)
        for (let i = 0; i < (buffer.length / 4); i++) {
            this.table[i] = buffer.readInt32LE(i * 4)
        }
    }

    getSecIdChain(startSecId) {
        var secId = startSecId;
        var secIds = [];
        while (secId != SEC_ID_END_OF_CHAIN) {
            secIds.push(secId);
            secId = this.table[secId];
        }
        return secIds;
    }
}

const ENTRY_TYPE_EMPTY = 0;
const ENTRY_TYPE_STORAGE = 1;
const ENTRY_TYPE_STREAM = 2;
const ENTRY_TYPE_ROOT = 5;
const NODE_COLOR_RED = 0;
const NODE_COLOR_BLACK = 1;
const LEAF = -1;

class DirectoryTree {
    /**
     * 
     * @param {OleCompoundDoc} doc 
     */
    constructor(doc) {
        this.doc = doc
    }

    async load(secIds) {
        let buffer = await this.doc.readSectors(secIds)
        let count = buffer.length / 128;
        /**
         * @type {StorageEntry[]}
         */
        this.entries = new Array(count);
        for (let i = 0; i < count; i++) {
            let offset = i * 128;

            let nameLength = Math.max(buffer.readInt16LE(64 + offset) - 1, 0);

            let entry = new StorageEntry({
                name: buffer.toString('utf16le', 0 + offset, nameLength + offset),
                type: buffer.readInt8(66 + offset),
                nodeColor: buffer.readInt8(67 + offset),
                left: buffer.readInt32LE(68 + offset),
                right: buffer.readInt32LE(72 + offset),
                storageDirId: buffer.readInt32LE(76 + offset),
                secId: buffer.readInt32LE(116 + offset),
                size: buffer.readInt32LE(120 + offset)
            });
            if (entry.type == ENTRY_TYPE_ROOT) {
                this.root = entry
            }
            this.entries[i] = entry;
        }
        this.buildHierarchy(this.root);
    }
    /**
     * 
     * @param {StorageEntry} storageEntry 
     */
    buildHierarchy(storageEntry) {
        let childIds = this.getChildIds(storageEntry)
        for (let i = 0; i < childIds.length; i++) {
            let childEntry = this.entries[childIds[i]]
            if (childEntry.type === ENTRY_TYPE_STORAGE) {
                storageEntry.storages[childEntry.name] = childEntry;
                this.buildHierarchy(childEntry)
            }
            if (childEntry.type === ENTRY_TYPE_STREAM) {
                storageEntry.streams[childEntry.name] = childEntry;
            }
        }
    }
    /**
     * 
     * @param {StorageEntry} storageEntry 
     */
    getChildIds(storageEntry) {
        let childs = []
        if (storageEntry.storageDirId > -1) {
            childs.push(storageEntry.storageDirId);
            var rootChildEntry = this.entries[storageEntry.storageDirId];
            rootChildEntry.visit(this.entries, childs)
        }
        return childs;
    }
}
class StorageEntry {
    constructor(entry) {
        this.name = entry.name
        this.type = entry.type
        this.nodeColor = entry.nodeColor
        this.left = entry.left
        this.right = entry.right
        this.storageDirId = entry.storageDirId
        this.secId = entry.secId
        this.size = entry.size
        this.storages = {}
        this.streams = {}
    }
    /**
     * 
     * @param {string} streamName 
     * @returns {StorageEntry}
     */
    stream(streamName) {
        return this.streams[streamName]
    }
    /**
     * 
     * @param {string} storageName 
     * @returns {StorageEntry}
     */
    storage(storageName) {
        return this.storages[storageName]
    }
    /**
     * 
     * @param {StorageEntry[]} entries 
     * @param {number[]} childIds 
     */
    visit(entries = [], childIds = []) {
        if (this.left !== LEAF) {
            childIds.push(this.left);
            entries[this.left].visit(entries, childIds)
        }
        if (this.right !== LEAF) {
            childIds.push(this.right);
            entries[this.right].visit(entries, childIds)
        }
    }
}


class OleCompoundDoc {
    /**
     * 
     * @param {string | Buffer} filename Name of a file or a Buffer with the data
     */
    constructor(filename) {
        if (filename instanceof Buffer) {
            this.data = filename
        } else {
            this.filename = filename;
        }

        this.skipBytes = 0;
    }

    async readWithCustomHeader(size) {
        this.skipBytes = size;

    }
    async read() {
        await this.openFile()
        await this.readHeader()
        await this.readMSAT()
        await this.readSAT()
        await this.readSSAT()
        await this.readDirectoryTree()
    }
    async openFile() {
        if (this.filename) {
            this.fd = await fs.open(this.filename, 'r');
        }
    }
    async readCustomHeader() {
        if (this.fd) {
            let buffer = Buffer.alloc(this.skipBytes)
            return (await this.fd.read(buffer, 0, this.skipBytes)).buffer
        } else {
            return this.data.slice(0, this.skipBytes)
        }
    }
    async readHeader() {

        if (this.fd) {
            let buffer = Buffer.alloc(512)
            await this.fd.read(buffer, 0, 512, this.skipBytes)
            this.header = new OleHeader();
            this.header.load(buffer)
        } else {
            this.header = new OleHeader();
            this.header.load(this.data.slice(this.skipBytes, this.skipBytes + 512))
        }

    }
    async readMSAT() {
        this.MSAT = this.header.partialMSAT.slice(0)
        this.MSAT.length = this.header.SATSize
        if (this.header.SATSize <= 109 || this.header.MSATSize == 0) {
            return
        }
        let buffer = Buffer.alloc(this.header.secSize)
        let currMSATIndex = 109;

        let secId = this.header.MSATSecId;
        for (let i = 0; i < this.header.MSATSize; i++) {
            let sectorBuffer = await this.readSector(secId)
            for (let s = 0; s < this.header.secSize - 4; s += 4) {
                if (currMSATIndex >= this.header.SATSize)
                    break;
                else
                    this.MSAT[currMSATIndex] = sectorBuffer.readInt32LE(s);

                currMSATIndex++;
            }
            secId = sectorBuffer.readInt32LE(this.header.secSize - 4);
        }
    }

    async readSector(secId) {
        return await this.readSectors([secId])
    }
    async readShortSector(secId) {
        return await this.readShortSectors([secId])
    }
    /**
     * 
     * @param {number[]} secIds 
     */
    async readShortSectors(secIds) {
        if (this.fd) {
            let buffer = Buffer.alloc(secIds.length * this.header.shortSecSize)
            for (let i = 0; i < secIds.length; i++) {
                let bufferOffset = i * this.header.shortSecSize;
                let fileOffset = this.getFileOffsetForShortSec(secIds[i])
                await this.fd.read(buffer, bufferOffset, this.header.shortSecSize, fileOffset)
            }
            return buffer
        } else {
            let buffers = []
            for (let i = 0; i < secIds.length; i++) {
                let fileOffset = this.getFileOffsetForShortSec(secIds[i])
                buffers.push(this.data.slice(fileOffset, fileOffset + this.header.shortSecSize))
            }
            return Buffer.concat(buffers)
        }
    }
    /**
     * 
     * @param {number[]} secIds 
     */
    async readSectors(secIds) {


        if (this.fd) {
            let buffer = Buffer.alloc(secIds.length * this.header.secSize)
            for (let i = 0; i < secIds.length; i++) {
                let bufferOffset = i * this.header.secSize;
                let fileOffset = this.getFileOffsetForSec(secIds[i])
                await this.fd.read(buffer, bufferOffset, this.header.secSize, fileOffset)
            }
            return buffer
        } else {
            let buffers = []
            for (let i = 0; i < secIds.length; i++) {
                let fileOffset = this.getFileOffsetForSec(secIds[i])
                buffers.push(this.data.slice(fileOffset, fileOffset + this.header.secSize))
            }
            return Buffer.concat(buffers)
        }
    }

    getFileOffsetForSec(secId) {
        return this.skipBytes + (secId + 1) * this.header.secSize;
    }
    getFileOffsetForShortSec(shortSecId) {
        var shortSecSize = this.header.shortSecSize;
        var shortStreamOffset = shortSecId * shortSecSize;

        var secIdIndex = Math.floor(shortStreamOffset / this.header.secSize);
        var secOffset = shortStreamOffset % this.header.secSize;
        var secId = this.shortStreamSecIds[secIdIndex];

        return this.getFileOffsetForSec(secId) + secOffset;
    }
    async readDirectoryTree() {
        this.directoryTree = new DirectoryTree(this)
        let secIds = this.SAT.getSecIdChain(this.header.dirSecId)
        await this.directoryTree.load(secIds)
        let rootEntry = this.directoryTree.root
        this.rootStorage = new Storage(this, rootEntry)
        this.shortStreamSecIds = this.SAT.getSecIdChain(rootEntry.secId)

    }
    /**
     * Read Sector Allocation Tablle
     */
    async readSAT() {
        this.SAT = new AllocationTable(this)
        await this.SAT.load(this.MSAT)
    }
    /**
     * Read Short Sector AllocationTable
     */
    async readSSAT() {
        this.SSAT = new AllocationTable(this)
        let secIds = this.SAT.getSecIdChain(this.header.SSATSecId)
        if (secIds.length != this.header.SSATSize) {
            throw new Error('Invalid Short Sector Allocation Table')
        }
        this.SSAT.load(secIds)
    }
    storage(storageName) {
        return this.rootStorage.storage(storageName)
    }
    stream(streamName) {
        return this.rootStorage.stream(streamName)
    }
    storageList() {
        return this.rootStorage.storageList()
    }
    streamList() {
        return this.rootStorage.streamList()
    }
}


class Storage {
    /**
     * 
     * @param {OleCompoundDoc} doc 
     * @param {StorageEntry} dirEntry 
     */
    constructor(doc, dirEntry) {
        this.doc = doc
        this.dirEntry = dirEntry
    }

    storageList() {
        return Object.keys(this.dirEntry.storages)
    }
    streamList() {
        return Object.keys(this.dirEntry.streams)
    }

    storage(storageName) {
        return new Storage(this.doc, this.dirEntry.storage(storageName))
    }
    async stream(streamName) {
        let streamEntry = this.dirEntry.stream(streamName);
        if (!streamEntry)
            return null
        let shortStream = false
        let allocationTable = this.doc.SAT
        let bytes = streamEntry.size;

        if (streamEntry.size < this.doc.header.shortStreamMax) {
            shortStream = true
            allocationTable = this.doc.SSAT
        }

        let secIds = allocationTable.getSecIdChain(streamEntry.secId)
        let retBuffer = []
        for (let i = 0; i < secIds.length; i++) {
            let buffer = shortStream ? await this.doc.readShortSector(secIds[i]) : await this.doc.readSector(secIds[i])
            if ((bytes - buffer.length) < 0) {
                buffer = buffer.slice(0, bytes);
            }

            bytes -= buffer.length;
            retBuffer.push(buffer)
        }


        return Buffer.concat(retBuffer)

    }
}

exports.OleCompoundDoc = OleCompoundDoc;