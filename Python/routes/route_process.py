from flask import Blueprint, request, jsonify, current_app
from prompts import get_prompt
import logging

process_callback = Blueprint('process', __name__)
logging.getLogger(__name__)

@process_callback.route('/process', methods=['GET', 'POST'])
def process():
    # 1. Extract arguments (add prompt_type)
    data = request.json if request.method == 'POST' else request.args
    text = data.get("text", "")
    task = data.get("task", "")
    prompt_type = data.get("prompt_type", "system_instruction") # Default to original
    
    print(f"[{request.remote_addr}] Request received: {task}")
    
    if not text:
        return "No text provided for processing", 400
    
    try:
        result = process_logic(text, task, prompt_type)
        return jsonify({"result": result})
    except Exception as e:
        err_msg = str(e).lower()
        print(f"LLM Response error: {err_msg}")
        status_code = 503 if "connection" in err_msg else 569
        return jsonify({"error": err_msg}), status_code
    
# In route_process.py
def process_logic(text, task, prompt_type="system_instruction"):
    client = current_app.config['OPENAI_CLIENT']
    model_name = current_app.config['MODEL_NAME']
    
    system_content = get_prompt(prompt_type)
    user_content = f"Task: {task}. Input: {text}"
    
    system_len = len(system_content)
    user_len = len(user_content)
    total_len = system_len + user_len
    
    print("DEBUG PROMPT LENGTHS:")
    print(f"    System: {system_len} chars (~{system_len//4} tokens)")
    print(f"    User:   {user_len} chars (~{user_len//4} tokens)")
    print(f"    Total:  {total_len} chars (~{total_len//4} tokens)")
    
    response = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": system_content},
            {"role": "user", "content": user_content}
        ]
    )
    return response.choices[0].message.content