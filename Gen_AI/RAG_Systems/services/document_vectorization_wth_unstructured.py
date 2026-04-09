from unstructured.partition.pdf import partition_pdf
from unstructured.chunking.title import chunk_by_title


from langchain_core.documents import Document

def partition_document(file_path: str):
    print(f"📄 Partitioning document: {file_path}")
    
    elements = partition_pdf(
        filename=file_path,
        strategy="hi_res",
        infer_table_structure=True,
        extract_image_block_types=["Image"],
        extract_image_block_to_payload=True
    )
    
    print(f"✅ Extracted {len(elements)} elements")
    return elements

def create_chunks_by_title(elements):
    """Create intelligent chunks using title-based strategy"""
    print("🔨 Creating smart chunks...")
    
    chunks = chunk_by_title(
        elements, # The parsed PDF elements from previous step
        max_characters=3000, # Hard limit - never exceed 3000 characters per chunk
        new_after_n_chars=2400, # Try to start a new chunk after 2400 characters
        combine_text_under_n_chars=500 # Merge tiny chunks under 500 chars with neighbors
    )
    
    print(f"✅ Created {len(chunks)} chunks")
    return chunks

def separate_content_types(chunk):
    """Analyze what types of content are in a chunk"""
    content_data = {
        'text': chunk.text,
        'tables': [],
        'images': [],
        'types': ['text']
    }
    
    # Check for tables and images in original elements
    if hasattr(chunk, 'metadata') and hasattr(chunk.metadata, 'orig_elements'):
        for element in chunk.metadata.orig_elements:
            element_type = type(element).__name__
            
            # Handle tables
            if element_type == 'Table':
                content_data['types'].append('table')
                table_html = getattr(element.metadata, 'text_as_html', element.text)
                content_data['tables'].append(table_html)
            
            # Handle images
            elif element_type == 'Image':
                if hasattr(element, 'metadata') and hasattr(element.metadata, 'image_base64'):
                    content_data['types'].append('image')
                    content_data['images'].append(element.metadata.image_base64)
    
    content_data['types'] = list(set(content_data['types']))
    return content_data

def process_elements(elements,generate_ai_summary_agent,claude_client):
    process_content=[]
    "Format Unstructured document elements for vectorization"
    image_cntr = 1
    for element in elements:
        element_type = type(element).__name__
        if element_type not in ['Table','Image']:
           process_content.append({
               'content':element.text,
               'metadata':{'last_modified':element.metadata.to_dict()['last_modified'],
                            'languages':element.metadata.to_dict()['languages'],
                            'page_number':element.metadata.to_dict()['page_number'],
                            'image_counter':None,
                            'image_base64':None,
                            'file_name':element.metadata.to_dict()['filename']}
           }) 
        elif element_type=='Table':
            table_html = getattr(element.metadata, 'text_as_html', element.text)
            soup = BeautifulSoup(table_html, "html.parser")
            pretty_html = soup.prettify()

            df = pd.read_html(StringIO(table_html))[0]
            table_text = df.to_string()

            process_content.append({
                'content':table_text,
                'metadata':{'last_modified':element.metadata.to_dict()['last_modified'],
                            'languages':element.metadata.to_dict()['languages'],
                            'page_number':element.metadata.to_dict()['page_number'],
                            'image_counter':None,
                            'image_base64':None,
                            'file_name':element.metadata.to_dict()['filename']}
                })
        elif element_type=='Image':
            image_summary = generate_ai_summary_agent(element.metadata.image_base64,claude_client)
            print(f"Image Summary:\n{image_summary}")
            process_content.append({
                'content':f"[Image-{image_cntr} Summary]\n\n{image_summary}",
                'metadata':{'last_modified':element.metadata.to_dict()['last_modified'],
                            'languages':element.metadata.to_dict()['languages'],
                            'page_number':element.metadata.to_dict()['page_number'],
                            'image_counter':f"[Image-{image_cntr}]",
                            'image_base64':element.metadata.image_base64,
                            'file_name':element.metadata.to_dict()['filename']}
            })
            image_cntr+=1
    return process_content