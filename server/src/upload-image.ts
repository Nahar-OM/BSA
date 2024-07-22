import { PutObjectCommand } from "@aws-sdk/client-s3";
import { getSignedUrl } from "@aws-sdk/s3-request-presigner";
import { Hono } from "hono";
import { s3Client } from "./config/s3-init";

const app = new Hono().post('/', async c => {
   try {
       const body = await c.req.parseBody({ all: true });
       const file = body.file;

       if (!file || !(file instanceof File)) {
           return c.json({ success: false, url: null, message: 'No file provided' }, 400);
       }

       // Ensure file is an .xlsx file
       const fileExtension = file.name.split('.').pop();
       if (fileExtension !== 'xlsx') {
           return c.json({ success: false, url: null, message: 'Invalid file type. Only .xlsx files are allowed.' }, 400);
       }

       const fileName = `${crypto.randomUUID()}.${fileExtension}`;
       const command = new PutObjectCommand({
           Bucket: process.env.AWS_S3_BUCKET_NAME,
           Key: fileName,
           ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
       });

       const signedUrl = await getSignedUrl(s3Client, command, { expiresIn: 3600 });

       const arrayBuffer = await file.arrayBuffer();
       await fetch(signedUrl, {
           method: 'PUT',
           body: arrayBuffer,
           headers: {
               'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
           },
       });

       const fileUrl = `https://${process.env.AWS_S3_BUCKET_NAME}.s3.amazonaws.com/${fileName}`;
       return c.json({ success: true, url: fileUrl, message: 'File uploaded to S3' }, 200);
   } catch (error) {
       console.error('Error uploading to S3:', error);
       return c.json({ success: false, url: null, message: 'Failed to upload file' }, 500);
   }
})

export default app;
