import type { IExecuteFunctions, INodeExecutionData, INodeType, INodeTypeDescription } from 'n8n-workflow';
import { NodeConnectionType, NodeOperationError } from 'n8n-workflow';
import { convertMarkdownToDocx } from '@mohtasham/md-to-docx';

export class MdToDocx implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Markdown to DOCX',
		name: 'mdToDocx',
		group: ['transform'],
		version: 1,
		description: 'Convert Markdown text to a DOCX document',
		defaults: {
			name: 'Markdown to DOCX',
		},
		inputs: [NodeConnectionType.Main],
		outputs: [NodeConnectionType.Main],
		usableAsTool: true,
		properties: [
			{
				displayName: 'Markdown Source',
				name: 'source',
				type: 'options',
				options: [
					{ name: 'From Field', value: 'fromField' },
					{ name: 'From Binary (Text File)', value: 'fromBinary' },
				],
				default: 'fromField',
				noDataExpression: true,
				description: 'Where to read the Markdown content from',
			},
			{
				displayName: 'Markdown Field',
				name: 'markdownField',
				type: 'string',
				default: 'markdown',
				required: true,
				displayOptions: {
					show: {
						source: ['fromField'],
					},
				},
				description: 'Name of the JSON field that contains the Markdown string',
			},
			{
				displayName: 'Binary Property',
				name: 'binaryPropertyName',
				type: 'string',
				default: 'data',
				required: true,
				displayOptions: {
					show: {
						source: ['fromBinary'],
					},
				},
				description: 'Name of the binary property containing a text/markdown file',
			},
			{
				displayName: 'Filename',
				name: 'filename',
				type: 'string',
				default: 'document.docx',
				required: true,
				description: 'Filename for the generated DOCX',
			},
			{
				displayName: 'Output Binary Property',
				name: 'outputBinaryProperty',
				type: 'string',
				default: 'data',
				required: true,
				description: 'Name of the binary property to store the DOCX file in',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnItems: INodeExecutionData[] = [];

		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			try {
				const source = this.getNodeParameter('source', itemIndex) as string;
				const filename = this.getNodeParameter('filename', itemIndex) as string;
				const outputBinaryProperty = this.getNodeParameter('outputBinaryProperty', itemIndex) as string;

				let markdown = '';
				if (source === 'fromField') {
					const fieldName = this.getNodeParameter('markdownField', itemIndex) as string;
					markdown = (items[itemIndex].json as Record<string, unknown>)[fieldName] as string;
					if (typeof markdown !== 'string') {
						throw new NodeOperationError(this.getNode(), `Field "${fieldName}" must be a string containing Markdown.`, { itemIndex });
					}
				} else {
					const binaryPropertyName = this.getNodeParameter('binaryPropertyName', itemIndex) as string;
					const binaryData = items[itemIndex].binary?.[binaryPropertyName];
					if (!binaryData) {
						throw new NodeOperationError(this.getNode(), `Binary property "${binaryPropertyName}" is missing.`, { itemIndex });
					}
					const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, binaryPropertyName);
					markdown = buffer.toString('utf8');
				}

				const blob = await convertMarkdownToDocx(markdown);
				// Convert Blob to Buffer (Node env)
				const arrayBuffer = await blob.arrayBuffer();
				const buffer = Buffer.from(arrayBuffer);

				const newItem: INodeExecutionData = {
					json: items[itemIndex].json,
					binary: {},
				};

				newItem.binary![outputBinaryProperty] = await this.helpers.prepareBinaryData(buffer, filename, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');

				returnItems.push(newItem);
			} catch (error) {
				if (this.continueOnFail()) {
					returnItems.push({ json: { error: (error as Error).message }, pairedItem: itemIndex });
					continue;
				}
				if (error.context) {
					error.context.itemIndex = itemIndex;
					throw error;
				}
				throw new NodeOperationError(this.getNode(), error, { itemIndex });
			}
		}

		return [returnItems];
	}
} 