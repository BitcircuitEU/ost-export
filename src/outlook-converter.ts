import * as fs from 'fs';
import * as path from 'path';
import { PSTFile, PSTFolder, PSTMessage } from 'pst-extractor';
import { exec } from 'child_process';
import { promisify } from 'util';

const execAsync = promisify(exec);
const writeFile = promisify(fs.writeFile);
const mkdir = promisify(fs.mkdir);
const readdir = promisify(fs.readdir);

// Disable console warnings and errors
const originalConsoleWarn = console.warn;
const originalConsoleError = console.error;
console.warn = () => {};
console.error = () => {};

interface MessageData {
    subject: string;
    sender: string;
    recipients: string[];
    body: string;
    sentDate?: Date;
    receivedDate?: Date;
    headers: string;
    attachments: Array<{
        filename: string;
        data: Buffer;
        contentType?: string;
    }>;
}

interface ContactData {
    fullName: string;
    email: string;
    businessPhone?: string;
    mobilePhone?: string;
    homePhone?: string;
    address?: string;
    company?: string;
    jobTitle?: string;
}

interface AppointmentData {
    subject: string;
    location: string;
    startTime?: Date;
    endTime?: Date;
    body: string;
}

interface FolderData {
    name: string;
    messages: MessageData[];
    contacts: ContactData[];
    appointments: AppointmentData[];
    subfolders: FolderData[];
}

function sanitizeFilename(filename: string): string {
    // First replace invalid characters
    let sanitized = filename.replace(/[<>:"/\\|?*]/g, '_');
    // Then truncate to maximum length (Windows has a 260 character path limit)
    // Leave room for file extension (4 chars) and some path depth
    return sanitized.slice(0, 180);
}

function formatDate(date?: Date): string {
    if (!date) return '';
    return date.toISOString().replace(/[:.]/g, '-');
}

function createEMLContent(message: MessageData): string {
    const boundary = 'boundary_' + Math.random().toString(36).substr(2);
    const date = message.sentDate || message.receivedDate || new Date();
    
    const headers = [
        'From: ' + message.sender,
        'To: ' + message.recipients.join(', '),
        'Subject: ' + message.subject,
        'Date: ' + date.toUTCString(),
        'MIME-Version: 1.0',
        message.attachments.length > 0 
            ? `Content-Type: multipart/mixed; boundary="${boundary}"`
            : (message.body.includes('<html') || message.body.includes('<body'))
                ? 'Content-Type: text/html; charset=utf-8'
                : 'Content-Type: text/plain; charset=utf-8',
        'X-Mailer: pst-extractor',
        message.headers || ''
    ].filter(Boolean).join('\r\n');

    if (message.attachments.length === 0) {
        return headers + '\r\n\r\n' + message.body;
    }

    const parts = [
        headers,
        '',
        `--${boundary}`,
        (message.body.includes('<html') || message.body.includes('<body'))
            ? 'Content-Type: text/html; charset=utf-8'
            : 'Content-Type: text/plain; charset=utf-8',
        '',
        message.body
    ];

    // Add attachments
    for (const attachment of message.attachments) {
        const filename = attachment.filename.replace(/"/g, '');
        parts.push(
            `--${boundary}`,
            `Content-Type: ${attachment.contentType || 'application/octet-stream'}`,
            `Content-Disposition: attachment; filename="${filename}"`,
            'Content-Transfer-Encoding: base64',
            '',
            attachment.data.toString('base64').match(/.{1,76}/g)?.join('\r\n') || ''
        );
    }

    parts.push(`--${boundary}--`);
    return parts.join('\r\n');
}

function createVCardContent(contact: ContactData): string {
    const vcard = [
        'BEGIN:VCARD',
        'VERSION:3.0',
        `FN:${contact.fullName}`,
        contact.email ? `EMAIL:${contact.email}` : '',
        contact.businessPhone ? `TEL;TYPE=WORK:${contact.businessPhone}` : '',
        contact.mobilePhone ? `TEL;TYPE=CELL:${contact.mobilePhone}` : '',
        contact.homePhone ? `TEL;TYPE=HOME:${contact.homePhone}` : '',
        contact.address ? `ADR:;;${contact.address}` : '',
        contact.company ? `ORG:${contact.company}` : '',
        contact.jobTitle ? `TITLE:${contact.jobTitle}` : '',
        'END:VCARD'
    ].filter(Boolean).join('\r\n');

    return vcard;
}

function createICalContent(appointment: AppointmentData): string {
    const ical = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'BEGIN:VEVENT',
        `SUMMARY:${appointment.subject}`,
        `LOCATION:${appointment.location || ''}`,
        appointment.startTime ? `DTSTART:${formatDate(appointment.startTime)}` : '',
        appointment.endTime ? `DTEND:${formatDate(appointment.endTime)}` : '',
        `DESCRIPTION:${appointment.body || ''}`,
        'END:VEVENT',
        'END:VCALENDAR'
    ].filter(Boolean).join('\r\n');

    return ical;
}

async function extractFolder(folder: PSTFolder): Promise<FolderData> {
    const folderData: FolderData = {
        name: folder.displayName || 'Unnamed Folder',
        messages: [],
        contacts: [],
        appointments: [],
        subfolders: []
    };

    if (folder.contentCount > 0) {
        try {
            let item = folder.getNextChild();
            while (item !== null) {
                try {
                    // Skip unknown message types
                    if (!item.messageClass?.startsWith('IPM.')) {
                        console.warn(`Skipping unknown message type: ${item.messageClass}`);
                        try {
                            item = folder.getNextChild();
                        } catch (nextError) {
                            console.error(`Error getting next item after unknown type: ${nextError}`);
                            break;
                        }
                        continue;
                    }

                    if (item.messageClass === 'IPM.Note' || item.messageClass.startsWith('IPM.Note.')) {
                        const message = item as PSTMessage;
                        const attachments = [];

                        // Extract attachments if present
                        if (message.numberOfAttachments > 0) {
                            for (let i = 0; i < message.numberOfAttachments; i++) {
                                try {
                                    const attachment = message.getAttachment(i);
                                    if (!attachment || !attachment.filename) {
                                        console.warn(`Skipping invalid attachment ${i}`);
                                        continue;
                                    }

                                    const attachmentStream = attachment.fileInputStream;
                                    if (!attachmentStream) {
                                        console.warn(`No stream available for attachment: ${attachment.filename}`);
                                        continue;
                                    }

                                    try {
                                        const chunks: Buffer[] = [];
                                        const bufferSize = 8176;
                                        const buffer = Buffer.alloc(bufferSize);
                                        let bytesRead: number;

                                        do {
                                            bytesRead = attachmentStream.read(buffer);
                                            if (bytesRead > 0) {
                                                chunks.push(Buffer.from(buffer.slice(0, bytesRead)));
                                            }
                                        } while (bytesRead === bufferSize);

                                        if (chunks.length > 0) {
                                            const attachmentBuffer = Buffer.concat(chunks);
                                            attachments.push({
                                                filename: sanitizeFilename(attachment.longFilename || attachment.filename),
                                                data: attachmentBuffer,
                                                contentType: attachment.mimeTag || 'application/octet-stream'
                                            });
                                        } else {
                                            console.warn(`No data read for attachment: ${attachment.filename}`);
                                        }
                                    } catch (streamError: any) {
                                        console.warn(`Error processing attachment stream for ${attachment.filename}: ${streamError}`);
                                        
                                        // Try alternative method for small attachments
                                        try {
                                            if (attachment.size && attachment.size < 1024 * 1024) { // < 1MB
                                                const rawData = [];
                                                let byte;
                                                while ((byte = attachmentStream.read()) !== -1) {
                                                    rawData.push(byte);
                                                }
                                                if (rawData.length > 0) {
                                                    attachments.push({
                                                        filename: sanitizeFilename(attachment.longFilename || attachment.filename),
                                                        data: Buffer.from(rawData),
                                                        contentType: attachment.mimeTag || 'application/octet-stream'
                                                    });
                                                }
                                            }
                                        } catch (fallbackError) {
                                            console.warn(`Fallback method failed for ${attachment.filename}: ${fallbackError}`);
                                        }
                                    }
                                } catch (attachError) {
                                    console.warn(`Error extracting attachment ${i}: ${attachError}`);
                                }
                            }
                        }

                        // Only add message if we could process it
                        try {
                            let emailBody = message.bodyHTML || message.body || '';
                            let bodyType = message.bodyHTML ? 'HTML' : 'Text';

                            // Handle RTF if present
                            if (message.bodyRTF) {
                                try {
                                    emailBody = message.bodyRTF.toString();
                                    bodyType = 'HTML';
                                } catch (rtfError) {
                                    console.warn(`Error converting RTF body: ${rtfError}`);
                                }
                            }

                            folderData.messages.push({
                                subject: sanitizeFilename(message.subject || '(No Subject)'),
                                sender: message.senderEmailAddress || '',
                                recipients: message.displayTo ? message.displayTo.split(';').map(r => r.trim()).filter(Boolean) : [],
                                body: emailBody,
                                sentDate: message.clientSubmitTime || message.messageDeliveryTime || undefined,
                                receivedDate: message.messageDeliveryTime || undefined,
                                headers: message.transportMessageHeaders || '',
                                attachments
                            });
                        } catch (messageError) {
                            console.warn(`Error adding message to folder data: ${messageError}`);
                        }
                    } else if (item.messageClass === 'IPM.Contact') {
                        folderData.contacts.push({
                            fullName: item.displayName || '',
                            email: '',  // Basic contact info only
                            businessPhone: '',
                            mobilePhone: '',
                            homePhone: '',
                            address: '',
                            company: '',
                            jobTitle: ''
                        });
                    } else if (item.messageClass === 'IPM.Appointment') {
                        folderData.appointments.push({
                            subject: item.subject || '',
                            location: '',  // Basic appointment info only
                            startTime: undefined,
                            endTime: undefined,
                            body: item.body || ''
                        });
                    }
                } catch (itemError) {
                    console.warn(`Error processing item: ${itemError}`);
                }

                try {
                    item = folder.getNextChild();
                } catch (nextError) {
                    console.warn(`Error getting next item: ${nextError}`);
                    break;
                }
            }
        } catch (error) {
            console.warn(`Error processing folder content: ${error}`);
        }
    }

    if (folder.hasSubfolders) {
        try {
            const childFolders = folder.getSubFolders();
            for (const childFolder of childFolders) {
                try {
                    const subfolderData = await extractFolder(childFolder);
                    if (subfolderData.messages.length > 0 || 
                        subfolderData.contacts.length > 0 || 
                        subfolderData.appointments.length > 0 || 
                        subfolderData.subfolders.length > 0) {
                        folderData.subfolders.push(subfolderData);
                    }
                } catch (subfolderError) {
                    console.warn(`Error processing subfolder: ${subfolderError}`);
                }
            }
        } catch (subfoldersError) {
            console.warn(`Error getting subfolders: ${subfoldersError}`);
        }
    }

    return folderData;
}

async function exportToFiles(data: FolderData, basePath: string): Promise<void> {
    const folderPath = path.join(basePath, sanitizeFilename(data.name));
    
    try {
        await mkdir(folderPath, { recursive: true });

        // Export messages
        if (data.messages.length > 0) {
            const emailsPath = path.join(folderPath, 'Emails');
            await mkdir(emailsPath, { recursive: true });
            
            for (const message of data.messages) {
                try {
                    const filename = sanitizeFilename(`${formatDate(message.sentDate)}_${message.subject}.eml`);
                    const emailPath = path.join(emailsPath, filename);
                    await writeFile(emailPath, createEMLContent(message));

                    // Save attachments separately
                    if (message.attachments.length > 0) {
                        const attachmentsPath = path.join(emailsPath, 'Attachments', sanitizeFilename(message.subject));
                        await mkdir(attachmentsPath, { recursive: true });
                        
                        for (const attachment of message.attachments) {
                            try {
                                const attachmentPath = path.join(attachmentsPath, sanitizeFilename(attachment.filename));
                                await writeFile(attachmentPath, attachment.data);
                            } catch (attachError) {
                                console.warn(`Error saving attachment: ${attachError}`);
                            }
                        }
                    }
                } catch (messageError) {
                    console.warn(`Error saving message: ${messageError}`);
                    continue;
                }
            }
        }

        // Export contacts
        if (data.contacts.length > 0) {
            const contactsPath = path.join(folderPath, 'Contacts');
            await mkdir(contactsPath, { recursive: true });
            
            for (const contact of data.contacts) {
                const filename = sanitizeFilename(`${contact.fullName}.vcf`);
                await writeFile(path.join(contactsPath, filename), createVCardContent(contact));
            }
        }

        // Export appointments
        if (data.appointments.length > 0) {
            const calendarPath = path.join(folderPath, 'Calendar');
            await mkdir(calendarPath, { recursive: true });
            
            for (const appointment of data.appointments) {
                const filename = sanitizeFilename(`${formatDate(appointment.startTime)}_${appointment.subject}.ics`);
                await writeFile(path.join(calendarPath, filename), createICalContent(appointment));
            }
        }

        // Process subfolders
        for (const subfolder of data.subfolders) {
            await exportToFiles(subfolder, folderPath);
        }
    } catch (error) {
        console.error(`Error exporting folder ${data.name}: ${error}`);
        throw error;
    }
}

async function convertOST(ostPath: string, outputPath: string): Promise<void> {
    try {
        if (!fs.existsSync(ostPath)) {
            throw new Error(`OST-Datei nicht gefunden: ${ostPath}`);
        }

        // Ensure output directory exists
        console.log(`Creating output directory: ${outputPath}`);
        await mkdir(outputPath, { recursive: true });

        console.log('Reading OST file...');
        const pstFile = new PSTFile(ostPath);
        
        console.log('Extracting folder structure...');
        const rootFolder = pstFile.getRootFolder();
        if (!rootFolder) {
            throw new Error('Konnte Root-Ordner nicht finden');
        }

        console.log('Processing folders and messages...');
        const folderData = await extractFolder(rootFolder);
        if (folderData.messages.length === 0 && 
            folderData.contacts.length === 0 && 
            folderData.appointments.length === 0 && 
            folderData.subfolders.length === 0) {
            throw new Error('Keine Daten gefunden');
        }
        
        console.log('Exporting to files...');
        await exportToFiles(folderData, outputPath);
        
        console.log('Export completed successfully!');
        console.log(`Output directory: ${outputPath}`);
    } catch (error) {
        console.error('Error during conversion:', error);
        throw error;
    }
}

async function findOSTFiles(directory: string): Promise<string[]> {
    const files = await readdir(directory);
    return files.filter(file => file.toLowerCase().endsWith('.ost'));
}

async function processAllOSTFiles() {
    try {
        // Use current working directory instead of __dirname
        const projectDir = process.cwd();
        const outputBaseDir = path.join(projectDir, 'outlook_export');
        
        // Create base output directory
        await mkdir(outputBaseDir, { recursive: true });

        // Find all OST files
        const ostFiles = await findOSTFiles(projectDir);
        
        if (ostFiles.length === 0) {
            console.log('Keine OST-Dateien im Projektordner gefunden.');
            return;
        }

        console.log(`${ostFiles.length} OST-Datei(en) gefunden. Starte Konvertierung...`);

        // Process each OST file
        for (const ostFile of ostFiles) {
            const ostPath = path.join(projectDir, ostFile);
            const outputDir = path.join(outputBaseDir, path.basename(ostFile, '.ost'));
            
            console.log(`\nVerarbeite: ${ostFile}`);
            try {
                await convertOST(ostPath, outputDir);
                console.log(`Konvertierung von ${ostFile} abgeschlossen.`);
            } catch (error) {
                // Continue with next file even if one fails
                console.log(`Fehler bei der Konvertierung von ${ostFile}`);
            }
        }

        console.log('\nAlle Konvertierungen abgeschlossen.');
    } catch (error) {
        console.log('Ein Fehler ist aufgetreten:', error);
    }
}

// Start processing
processAllOSTFiles()
    .then(() => {
        console.log('Drücken Sie eine beliebige Taste zum Beenden...');
        process.stdin.setRawMode(true);
        process.stdin.resume();
        process.stdin.on('data', process.exit.bind(process, 0));
    })
    .catch(() => process.exit(1)); 