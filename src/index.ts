import * as fs from 'fs';
import * as path from 'path';
import { PSTFile, PSTFolder, PSTMessage } from 'pst-extractor';

function processFolder(folder: PSTFolder, depth: number = 0): void {
    const indent = '  '.repeat(depth);
    
    // Print current folder
    if (depth > 0) { // Skip root folder name
        console.log(`${indent}📁 ${folder.displayName}`);
    }

    // Process messages in current folder
    if (folder.contentCount > 0) {
        try {
            let email: PSTMessage | null = folder.getNextChild();
            while (email !== null) {
                try {
                    const subject = email.subject || '(No Subject)';
                    console.log(`${indent}  📧 ${subject}`);
                } catch (messageError) {
                    console.error(`Error processing message: ${messageError}`);
                }
                try {
                    email = folder.getNextChild();
                } catch (nextError) {
                    console.error(`Error getting next message: ${nextError}`);
                    break;
                }
            }
        } catch (contentError) {
            console.error(`Error processing folder content: ${contentError}`);
        }
    }

    // Process subfolders
    if (folder.hasSubfolders) {
        try {
            const childFolders = folder.getSubFolders();
            for (const childFolder of childFolders) {
                try {
                    processFolder(childFolder, depth + 1);
                } catch (subfolderError) {
                    console.error(`Error processing subfolder: ${subfolderError}`);
                }
            }
        } catch (subfoldersError) {
            console.error(`Error getting subfolders: ${subfoldersError}`);
        }
    }
}

async function readOSTFile(ostFilePath: string): Promise<void> {
    try {
        // Check if input file exists
        if (!fs.existsSync(ostFilePath)) {
            throw new Error(`OST file not found at path: ${ostFilePath}`);
        }

        console.log('Starting to read OST file...');
        console.log(`Input file: ${ostFilePath}`);

        // Attempt to open the OST file
        const pstFile = new PSTFile(ostFilePath);
        console.log('\nFile Details:');
        console.log('-------------');
        console.log(`Message Store: ${pstFile.getMessageStore().displayName}`);
        
        console.log('\nFolder Structure:');
        console.log('----------------');
        // Process the root folder and all its subfolders
        processFolder(pstFile.getRootFolder());

        console.log('\nFile reading completed successfully!');
    } catch (error) {
        console.error('Error while reading OST file:', error);
        throw error;
    }
}

// Example usage
const inputFile = path.join(__dirname, '..', 'niklas@terhorst.io.ost');

readOSTFile(inputFile)
    .then(() => console.log('Process completed'))
    .catch((error) => console.error('Process failed:', error)); 