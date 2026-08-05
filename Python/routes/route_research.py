from flask import Blueprint, request, jsonify, current_app
from .route_process import process_logic
from urllib.parse import quote
import requests, logging, re, json
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
    top_k = data.get("topK", "4")
    
    if not search_term:
        return jsonify({"status": "error", "code": 400, "message": "Search term is empty"}), 400

    # 2. Query Optimization (via process_logic)
    try:
        optimized_queries_raw = process_logic(
            text=search_term,
            task="Generate distinct (Minimum 3 - Max 5), specific search queries. Return only queries, one per line.",
            prompt_type="system_instruction_search_query"
        )
        queries = [q.strip() for q in optimized_queries_raw.split('\n') if q.strip()]
        print("\nDEBUG Queries:")
        print("\n".join(f"    {q}" for q in queries))
        
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
    print("\nDEBUG Zotseek search results:")
    print(f"    Obtained {len(all_results)} results from ZotSeek")
    grouped_results = {}
    total_processed = 0
    total_duplicates = 0
    seen_pairs = {}

    # First, group all results by itemKey
    for idx, item in enumerate(all_results, 1):
        snippet = item.get('matchedChunk', {}).get('snippet', '')
        item_key = item.get('itemKey', '')
        page = item.get('matchedChunk', {}).get('page', None)
        title = item.get('title', '')
        authors = item.get('authors', '')
        year = item.get('year', '')
        
        # Clean up extra spaces
        snippet = re.sub(r'\s+', ' ', snippet).strip()
        
        # Exact dedup: same itemKey + same page
        pair_key = f"{item_key}_{page}"
        if pair_key in seen_pairs:
            total_duplicates += 1
            if total_duplicates == 1:
                print(f"    Grouping duplicate results by (itemKey, page)")
            first_idx = seen_pairs[pair_key]
            print(f"        >> Result {first_idx} (page {page}) and Result {idx} (page {page}) from {item_key}")
            continue
        seen_pairs[pair_key] = idx
        
        total_processed += 1
        
        # Group by itemKey
        if item_key not in grouped_results:
            grouped_results[item_key] = {
                'snippets': [],
                'ref_nums': [],
                'pages': [],
                'title': title,
                'authors': authors,
                'year': year
            }
        
        grouped_results[item_key]['snippets'].append(snippet)
        grouped_results[item_key]['ref_nums'].append(idx)
        grouped_results[item_key]['pages'].append(page)
                        
    # Now build context string with grouped snippets
    ref_id_counter = 1
    context_list = []  # Store JSON contexts for each item
    
    print("\nDEBUG Search context:")

    for item_key, data in grouped_results.items():
        # Build JSON context for LLM first
        json_context = {
            "source": {
                "citation_format": f"zotero://{item_key}",
                "authors": data['authors'],
                "title": data['title'],
                "year": data['year'],
                "total_snippets": len(data['snippets'])
            },
            "snippets": [
                {
                    "content": f"Page {data['pages'][i]}: {data['snippets'][i]}"
                }
                for i in range(len(data['snippets']))
                if data['snippets'][i] is not None  # optional: filter out None snippets
            ]
        }
        context_list.append(json_context)
        
        # Print debug info from the JSON structure
        print(f"    RAG Context: {ref_id_counter}")
        print(f"    zotero_itemKey: {item_key}")
        print(f"    Authors: {json_context['source']['authors']}")
        title = json_context['source']['title']
        print(f"    Title: {title[:80]}..." if len(title) > 80 else f"    Title: {title}")
        print(f"    Year: {json_context['source']['year']}")
        print(f"    Appears in {json_context['source']['total_snippets']} snippets (ref_nums: {data['ref_nums']})")
        
        for snippet_item in json_context['snippets']:
            snippet_text = snippet_item['content']
            if len(snippet_text) > 80:
                print(f"        snippet: {snippet_text[:80]}...\n")
            else:
                print(f"        snippet: {snippet_text}\n")
        
        ref_id_counter += 1

    # Combine all JSON contexts with clear separation
    context_str = json.dumps(context_list, indent=2)
    
    print(f"\nTotal results: {len(all_results)}")
    print(f"Processed: {total_processed}")
    print(f"Duplicates removed (same itemKey + page): {total_duplicates}")
    print(f"Unique groups: {len(grouped_results)}")

    try:
        synthesized_text = process_logic(
            text=context_str,
            task=f"User request/context: {search_term}",
            prompt_type="system_instruction_research"
        )
        return synthesized_text, 200, {'Content-Type': 'text/plain'}
    except Exception as e:
        error_msg = f"Synthesis failed: {str(e)}"
        print(f"Error: {error_msg}")
        return error_msg, 569