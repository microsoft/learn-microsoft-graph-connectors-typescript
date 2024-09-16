import { GraphError } from '@microsoft/microsoft-graph-client';
import { ExternalConnectors } from '@microsoft/microsoft-graph-types';
import { config } from './config.js';
import { client } from './graphClient.js';
import matter, { GrayMatterFile } from 'gray-matter';
import fs from 'fs';
import path from 'path';
import removeMd from 'remove-markdown';

// Represents the document to import
interface Document extends GrayMatterFile<string> {
  content: string;
  relativePath: string;
  url?: string;
  iconUrl?: string;
}

function extract(): Document[] {
  const contentDir = 'content';
  const baseUrl = 'https://learn.microsoft.com/graph/';

  const content: Document[] = [];
  const contentFiles = fs.readdirSync(contentDir, { recursive: true });

  contentFiles.forEach(file => {
    if (!file.toString().endsWith('.md')) {
      return;
    }

    const fileContents = fs.readFileSync(path.join(contentDir, file.toString()), 'utf-8');
    const doc = matter(fileContents) as Document;

    doc.content = removeMd(doc.content.replace(/<[^>]+>/g, ' '));
    doc.relativePath = file.toString();
    doc.url = new URL(doc.relativePath.replace('.md', ''), baseUrl).toString();
    doc.iconUrl = 'https://raw.githubusercontent.com/waldekmastykarz/img/main/microsoft-graph.png';

    content.push(doc);
  });

  return content;
}

function getDocId(doc: Document): string {
  const id = doc.relativePath.replace(path.sep, '__').replace('.md', '');
  return id;
}

function transform(documents: Document[]): ExternalConnectors.ExternalItem[] {
  return documents.map(doc => {
    const docId = getDocId(doc);

    let acl = [
      {
          accessType: 'grant',
          type: 'everyone',
          value: 'everyone'
        }
    ];

    if (doc.relativePath.endsWith('use-the-api.md')) {
      acl = [
        {
          accessType: 'grant',
          type: 'user',
          value: '2e75bd61-7a32-44aa-b8a7-ff051804df25'
        },
      ];
    }
    else if (doc.relativePath.endsWith('traverse-the-graph.md')) {
      acl = [ 
        {
          accessType: 'grant',
          type: 'group',
          value: 'a9fd282f-4634-4cba-9dd4-631a2ee83cd3',
        }
      ];
    }

    return {
      id: docId,
      properties: {
        title: doc.data.title ?? '',
        description: doc.data.description ?? '',
        url: doc.url,
        iconUrl: doc.iconUrl
      },
      content: {
        value: doc.content ?? '',
        type: 'text'
      },
      acl: acl,
    } as ExternalConnectors.ExternalItem
  });
}

async function load(externalItems: ExternalConnectors.ExternalItem[]) {
  const { id } = config.connection;
  for (const doc of externalItems) {
    try {
      console.log(`Loading ${doc.id}...`);
      await client
        .api(`/external/connections/${id}/items/${doc.id}`)
        .header('content-type', 'application/json')
        .put(doc);
      console.log('  DONE');
    }
    catch (e) {
      const graphError = e as GraphError;
      console.error(`Failed to load ${doc.id}: ${graphError.message}`);
      if (graphError.body) {
        console.error(`${JSON.parse(graphError.body)?.innerError?.message}`);
      }
      return;
    }
  }
}

export async function loadContent() {
  const content = extract();
  const transformed = transform(content);
  await load(transformed);
}

loadContent();