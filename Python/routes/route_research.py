from flask import Blueprint, request, jsonify, current_app
from .route_process import process_logic
from urllib.parse import quote
import requests, logging, re
from difflib import SequenceMatcher

research_callback = Blueprint('research', __name__)
logging.getLogger(__name__)

# ZotSeek error mappings
ZOTSEEK_ERROR_MAPPINGS = {
    400: "Invalid search request. Check your syntax.",
    404: "Zotseek AI Agent API disabled. Please enable it!",
    500: "Zotero internal server error.",
    503: "Zotero appears to be offline. Is the local API running?"
}

@research_callback.route('/research', methods=['GET', 'POST'])
def research():
    # 1. Extract arguments
    data = request.json if request.method == 'POST' else request.args
    search_term = data.get("text", "")
    top_k = data.get("topK", "3")
    
    if not search_term:
        return jsonify({"status": "error", "code": 400, "message": "Search term is empty"}), 400

    # 2. Query Optimization (via process_logic)
    try:
        optimized_queries_raw = process_logic(
            text=search_term,
            task="Generate 3 distinct, specific search queries. Return only queries, one per line.",
            prompt_type="system_instruction_search_query"
        )
        queries = [q.strip() for q in optimized_queries_raw.split('\n') if q.strip()]
        print("DEBUG Queries:")
        print("\n".join(f"\t{q}" for q in queries))
        
    except Exception as e:
        error_msg = f"LLM Query Optimization failed: {str(e)}"
        print(f"Error: {error_msg}")
        return error_msg, 569

    # queries = [search_term]
    
    # 3. ZotSeek Search with error handling
    all_results = []
    for q in queries:
        encoded_search = quote(q)
        url = f"http://localhost:23119/zotseek/search?q={encoded_search}&topK={top_k}"
        try:
            resp = requests.get(url, timeout=30)
            if resp.status_code == 200:
                all_results.extend(resp.json().get("results", []))
            elif resp.status_code in ZOTSEEK_ERROR_MAPPINGS:
                error_msg = ZOTSEEK_ERROR_MAPPINGS.get(resp.status_code, "Zotero error")
                print(f"Error code: {resp.status_code}; Error in Request: {error_msg}")
                return error_msg, resp.status_code
        except requests.exceptions.ConnectionError:
            error_msg = ZOTSEEK_ERROR_MAPPINGS.get(503, "Could not connect to Zotero Local API.")
            print(f"Error code: 503; Error in Request: {error_msg}")
            return error_msg, 503
        except Exception as e:
            error_msg = f"Unknown error from Zotseek: {str(e)}"
            print(f"Error in Request from Zotero: {error_msg}")
            return error_msg, 569

    if not all_results:
        error_msg = "No results found from Zotero."
        print(f"Error: {error_msg}")
        return error_msg, 404

    # 4. Synthesize - Group by itemKey with clear separation
    print("DEBUG Search context:")
    context_str = ""
    grouped_results = {}

    # First, group all results by itemKey
    for idx, item in enumerate(all_results, 1):
        matched_chunk = item.get('matchedChunk')
        
        # Skip if matchedChunk is None
        if matched_chunk is None:
            continue
        
        snippet = item.get('matchedChunk', {}).get('snippet', '')
        item_key = item.get('itemKey', '')
        
        # Remove any [ref_num: ...] or [number, itemKey] patterns
        snippet = re.sub(r'\[ref_num:\s*\d+\s*,\s*zotero_itemKey:\s*[A-Z0-9]+\]', '', snippet)
        snippet = re.sub(r'\[\d+\s*,\s*[A-Z0-9]+\]', '', snippet)
        snippet = re.sub(r'\[\d+\]', '', snippet)
        
        # Clean up extra spaces
        snippet = re.sub(r'\s+', ' ', snippet).strip()
        
        # Group by itemKey
        if item_key not in grouped_results:
            grouped_results[item_key] = {
                'snippets': [],
                'ref_nums': []
            }
            
        # Check if the snippet is a duplicate of any already stored snippet for this itemKey
        is_duplicate = False
        for existing_snippet in grouped_results[item_key]['snippets']:
            if is_similar(snippet, existing_snippet, threshold=0.85):
                is_duplicate = True
                break
                
        if not is_duplicate:
            grouped_results[item_key]['snippets'].append(snippet)
            grouped_results[item_key]['ref_nums'].append(idx)

    # Now build context string with grouped snippets
    ref_id_counter = 1
    for item_key, data in grouped_results.items():
        # Build context string - show each snippet individually
        # context_str += f"--- SOURCE REFERENCE ID: {ref_id_counter} ---\n"
        context_str += f"--- RAG Context: {ref_id_counter} ---\n"
        context_str += f"zotero_itemKey: {item_key}\n"
        context_str += f"This source appears in {len(data['snippets'])} different paragraphs (ref_nums: {', '.join(map(str, data['ref_nums']))})\n"
        
        for i, snippet in enumerate(data['snippets'], 1):
            context_str += f"  Content {i}: {snippet}\n"
        
        context_str += "\n"
        
        # Print debug info
        print(f"\tRAG Context: {ref_id_counter}")
        print(f"\tzotero_itemKey: {item_key}")
        print(f"\tAppears in {len(data['snippets'])} snippets (ref_nums: {data['ref_nums']})")
        for i, snippet in enumerate(data['snippets'], 1):
            print(f"\t\tContent {i}: {snippet[:80]}..." if len(snippet) > 80 else f"\t\tContent {i}: {snippet}")
        print()
        
        ref_id_counter += 1
    
    try:
        synthesized_text = process_logic(
            text=context_str,
            task=f"User prompt/context: {search_term}. Remember the citation format 'zotero://XXXXXXXX'",
            prompt_type="system_instruction_research"
        )
        # return jsonify({
        #     "status": "success", 
        #     "raw": {"results": all_results}, 
        #     "synthesized": synthesized_text
        # })
        return synthesized_text, 200, {'Content-Type': 'text/plain'}
    except Exception as e:
        error_msg = f"Synthesis failed: {str(e)}"
        print(f"Error: {error_msg}")
        return error_msg, 569
    
def is_similar(a: str, b: str, threshold: float = 0.85) -> bool:
    # Normalize whitespace
    a_norm = ' '.join(a.strip().split())
    b_norm = ' '.join(b.strip().split())
    
    ratio = SequenceMatcher(None, a_norm, b_norm).ratio()
    
    # # Diagnostic printing
    # print(f"  Similarity check:")
    # print(f"    A ({len(a_norm)} chars): {a_norm[:150]}{'...' if len(a_norm) > 150 else ''}")
    # print(f"    B ({len(b_norm)} chars): {b_norm[:150]}{'...' if len(b_norm) > 150 else ''}")
    # print(f"    Ratio: {ratio:.4f} (threshold: {threshold})")
    # print(f"    Result: {'DUPLICATE' if ratio >= threshold else 'UNIQUE'}")
    # print()
    
    return ratio >= threshold