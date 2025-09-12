import type { IExecuteFunctions, INodeExecutionData, INodeType, INodeTypeDescription } from 'n8n-workflow';
import { NodeConnectionType, NodeOperationError } from 'n8n-workflow';
import { convertMarkdownToDocx } from '@mohtasham/md-to-docx';

export class MdToDocx implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Markdown to DOCX',
		name: 'mdToDocx',
		icon: 'file:mdToDocx.svg',
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
			{
				displayName: 'Document Type',
				name: 'documentType',
				type: 'options',
				options: [
					{ name: 'Document', value: 'document' },
					{ name: 'Report', value: 'report' },
				],
				default: 'document',
				description: 'Type of document to generate',
			},
			{
				displayName: 'Advanced Options',
				name: 'advancedOptions',
				type: 'collection',
				placeholder: 'Add Option',
				default: {},
				options: [
					{
						displayName: 'Heading 1 Size',
						name: 'heading1Size',
						type: 'number',
						default: 48,
						description: 'Font size for H1 headings (in half-points, so 48 = 24pt)',
					},
					{
						displayName: 'Heading 2 Size',
						name: 'heading2Size',
						type: 'number',
						default: 36,
						description: 'Font size for H2 headings (in half-points, so 36 = 18pt)',
					},
					{
						displayName: 'Heading Spacing',
						name: 'headingSpacing',
						type: 'number',
						default: 240,
						description: 'Spacing before/after headings',
					},
					{
						displayName: 'Line Spacing',
						name: 'lineSpacing',
						type: 'number',
						default: 1.15,
						description: 'Line spacing multiplier',
					},
					{
						displayName: 'Paragraph Alignment',
						name: 'paragraphAlignment',
						type: 'options',
						options: [
							{ name: 'Left', value: 'LEFT' },
							{ name: 'Right', value: 'RIGHT' },
							{ name: 'Center', value: 'CENTER' },
							{ name: 'Justified', value: 'JUSTIFIED' },
						],
						default: 'LEFT',
						description: 'Text alignment for paragraphs',
					},
					{
						displayName: 'Paragraph Size',
						name: 'paragraphSize',
						type: 'number',
						default: 24,
						description: 'Font size for paragraphs (in half-points, so 24 = 12pt)',
					},
					{
						displayName: 'Paragraph Spacing',
						name: 'paragraphSpacing',
						type: 'number',
						default: 240,
						description: 'Spacing before/after paragraphs',
					},
					{
						displayName: 'Text Direction',
						name: 'direction',
						type: 'options',
						options: [
							{ name: 'Left to Right (LTR)', value: 'LTR' },
							{ name: 'Right to Left (RTL)', value: 'RTL' },
						],
						default: 'LTR',
						description: 'Text direction for the document (useful for Arabic, Hebrew, etc.)',
					},
					{
						displayName: 'Title Size',
						name: 'titleSize',
						type: 'number',
						default: 48,
						description: 'Font size for document title (in half-points, so 48 = 24pt)',
					},
				],
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
				const documentType = this.getNodeParameter('documentType', itemIndex) as string;
				const advancedOptions = this.getNodeParameter('advancedOptions', itemIndex) as Record<string, any>;

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

				// Convert literal \n characters to actual newlines
				markdown = markdown.replace(/\\n/g, '\n');

				// Build options object for convertMarkdownToDocx
				const options: any = {
					documentType,
				};

				// Always provide style options with defaults to ensure headers are formatted
				// Note: Font sizes are in half-points (so 32 = 16pt, 24 = 12pt, etc.)
				options.style = {
					// Default sizes that ensure headers are properly formatted
					titleSize: 48,        // 24pt for main title
					heading1Size: 48,     // 24pt for H1
					heading2Size: 36,     // 18pt for H2  
					heading3Size: 32,     // 16pt for H3
					heading4Size: 28,     // 14pt for H4
					heading5Size: 24,     // 12pt for H5
					paragraphSize: 24,    // 12pt for paragraphs
					listItemSize: 24,     // 12pt for list items
					codeBlockSize: 20,    // 10pt for code blocks
					blockquoteSize: 24,   // 12pt for blockquotes
					headingSpacing: 240,  // Spacing before/after headings
					paragraphSpacing: 240, // Spacing before/after paragraphs
					lineSpacing: 1.15,    // Line spacing multiplier
					paragraphAlignment: 'LEFT',
					direction: 'LTR',     // Default text direction
					// Override with any user-provided advanced options
					...advancedOptions,
				};

				const blob = await convertMarkdownToDocx(markdown, options);
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