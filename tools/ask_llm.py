
import sys
if "." not in sys.path:
    sys.path.append(".")
from copy import deepcopy

from megfile import smart_open
import base64

import openai
import os

# from tools.video_tools import extra_less_than_given_frames
import json
import time

def find_model_privider_from_name(model_name):
    if "gpt-" in model_name or model_name in ["Qwen2-VL-72B"]:
        return "openai"
    elif "step-1v" in model_name:
        return "step_official"
    elif model_name in ['qwen-vl-max-0809']:
        return "qwen_official"
    else:
        return "step"
    
    
from tqdm import tqdm
import requests


def directly_ask_gemini(url, api_key, model, messages, args={}):

    headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            }
    data = {
        "model": model,
        "stream": False,
        "messages": messages,
        "max_tokens": args.get("max_tokens", 100),
        "temperature": args.get("temperature", 0.5),
        "top_p": args.get("top_p", 1.0),
        "frequency_penalty": args.get("frequency_penalty", 0.0),

    }
    # print(data)
    response = requests.post(url + "/chat/completions", headers=headers, data=json.dumps(data), timeout=200)
    elapsed_time = response.elapsed.total_seconds()
    # print(f'model:{model}, status:{response.status_code}, cost:{elapsed_time}')
    # if response.status_code != 200:
        # print(f'response error. {response.status_code} error:{response.text}')
        # return "response error"
    # print(response.content.decode("utf-8"))
    answer = response.json()
    return answer


def ask_llm_anything(model_provider, model_name, messages, args={
    "max_tokens": 256,
    "temperature": 0.5,
    "top_p": 1.0,
    "frequency_penalty": 0.0,
}, tools=None, return_origin_response = False, use_stream = False, quiet = True, use_image_url = True):

    if model_provider is None:
        model_provider = "step"

    if not quiet:
        print(f"model_provider:{model_provider}, model_name:{model_name}")
    if model_provider != "step" or use_image_url:
        # to force convert the images in massage to base64 format
        for msg in messages:
            if type(msg['content']) == str:
                continue
            assert type(msg['content']) == list
            for content in msg['content']:
                if content['type'] == "text":
                    continue
                assert content['type'] == "image_url" or content['type'] == "image_b64"
                if content['type'] == "image_url":
                    url = content['image_url']['url']
                    # to check if the image is already in base64 format
                    if url.startswith("data:image/"):
                        continue
                    else:
                        image_bytes = smart_open(url, mode="rb").read()
                        b64 = base64.b64encode(image_bytes).decode('utf-8')

                        # to judge the image format
                        if image_bytes[0:4] == b"\x89PNG":
                            content['image_url']['url'] = "data:image/png;base64," + b64
                        elif image_bytes[0:2] == b"\xff\xd8":
                            content['image_url']['url'] = "data:image/jpeg;base64," + b64
                        else:
                            content['image_url']['url'] = "data:image/png;base64," + b64


                        # b64 = base64.b64encode(image_bytes).decode('utf-8')
                        # content['image_url']['url'] = "data:image/png;base64," + b64
                    
                else:
                    assert content['type'] == "image_b64"
                    b64 = content['image_b64']['b64_json']
                    del content['image_b64']
                    content['image_url'] = {"url": "data:image/png;base64," + b64}
                    content['type'] = "image_url"
        # start to ask openai 
    elif "step" == model_provider.lower():
        openai.api_key = "EMPTY"
        # openai.api_base = 'http://stepcast-router.basemind-core.svc.platform.basemind.local:9200/v1'
        # openai.api_base = "http://stepcast-router:9200/v1"
        openai.api_base = "http://10.148.42.15:9200/v1"
        # to change image_url to image_b64
        for msg in messages:
            if type(msg['content']) == str:
                continue
            assert type(msg['content']) == list
            for content in msg['content']:
                if content['type'] == "text":
                    continue
                assert content['type'] == "image_url" or content['type'] == "image_b64"
                # content['detail'] = "low"
                # content['detail'] = "high"
                if content['type'] == "image_url":
                    url = content['image_url']['url']
                    # print(url, flush=True)
                    
                    # to check if the image is already in base64 format
                    if url.startswith("data:image/"):
                        b64 = url.split(",")[1]
                        content['image_b64'] = {"b64": b64}
                        del content['image_url']
                        content['type'] = "image_b64"
                    else:
                        image_bytes = smart_open(url, mode="rb").read()
                        
                        b64 = base64.b64encode(image_bytes).decode('utf-8')
                        content['image_b64'] = {"b64_json": b64}
                        del content['image_url']
                        content['type'] = "image_b64"
                else:
                    assert content['type'] == "image_b64"


    for msg in messages:
        # print(msg['role'],":",end=" ")
        if msg['role'] == "human":
            msg['role'] = "user"
        prompts = []
        if type(msg['content']) == str:
            # print(msg['content'])
            prompts.append(msg['content'])
        else:
            for content in msg['content']:
                if content['type'] == "text":
                    prompts.append(content['text'])
                    # print(content['text'])
                elif "image" in content['type']:
                    prompts.append("<image>")
        # print("".join(prompts))

    http_proxy = os.environ.get('http_proxy', "")

    os.environ['https_proxy'] = ""
    if model_provider == "step":

        os.environ['https_proxy'] = ""
        openai.api_key = "EMPTY"
        # openai.api_base = "http://stepcast-router:9200/v1"
        # openai.api_base = "http://10.148.42.15:9200/v1"
        openai.api_base = "https://stepcast-router-eval-common.stepfun-inc.com/v1"


    elif model_provider == "openai":
        os.environ['https_proxy'] = ""

        openai.api_base = 'https://models-proxy.stepfun-inc.com/v1'
        #openai.api_key = "ak-68d2efa11e2ab28ccac3e3e6c825b6a1"
        openai.api_key = "ak-64c9efbg17h3jkl82mno59pqrs43tuv0w6x2z5"

    elif model_provider == "step_official":
        os.environ['https_proxy'] = "http://proxy.i.basemind.com:3128"

        openai.api_base = 'https://api.stepfun.com/v1'
        openai.api_key = "8Ffve6Mpu2WWmVswFcn9HJoNPPtDRtdr3PONJme4V0YWETYKZTSNvyTKO7UDbXnk"

    elif model_provider == "qwen_official":
        os.environ['https_proxy'] = "http://proxy.i.basemind.com:3128"

        openai.api_base = "https://dashscope.aliyuncs.com/compatible-mode/v1"
        openai.api_key = "sk-f250e914b5ee4d61a120bf6e8ac7f8b6"

    elif model_provider == "t":
        os.environ['https_proxy'] = ""
        openai.api_key = "EMPTY"

        openai.api_base = "http://100.96.159.21:8000/v1"
        # openai.api_key = "sk-f250e914b5ee4d61a120bf6e8ac7f8b6"    


    elif model_provider == "step_test":

        # https://api.c.ibasemind.com/v1/
        # 431YzzM6IyZ7GSuZbBXIzQFoaBMQy8T7NXlUzgWAUfFDDt2AbNZq9UBYLL9tfgOzt
        os.environ['https_proxy'] = ""
        openai.api_key = "431YzzM6IyZ7GSuZbBXIzQFoaBMQy8T7NXlUzgWAUfFDDt2AbNZq9UBYLL9tfgOzt"
        openai.api_base = "https://api.c.ibasemind.com/v1"

    elif model_provider.startswith("http"):
        os.environ['https_proxy'] = ""
        openai.api_key = "EMPTY"
        openai.api_base = model_provider
    
    elif model_provider == "gemini":
        os.environ['https_proxy'] = ""
        api_key = "ak-68d2efa11e2ab28ccac3e3e6c825b6a1"
        base_url = "https://models-proxy.stepfun-inc.com/v1"
        model = model_name
        complation = directly_ask_gemini(base_url, api_key, model, messages, args)
        answer = complation['choices'][0]['message']['content']
        if return_origin_response:
            return answer, complation
        
        return answer

    os.environ["https_proxy"] = os.environ["http_proxy"] = ""
    completion = openai.ChatCompletion.create(
        api_key=openai.api_key,
        api_base = openai.api_base,
        model=model_name,
        messages=messages,
        temperature=args.get("temperature", 0.5),
        top_p=args.get("top_p", 1.0),
        frequency_penalty=args.get("frequency_penalty", 0.0),
        max_tokens=args.get("max_tokens", 100), 
        tools = tools,
        ## tmp
        # thinking = {"type": "enabled"}
    )

    
    os.environ["https_proxy"] = os.environ["http_proxy"] = http_proxy

    def show_completion(completion):
        full_response = ""
        for chunk in completion:
            if chunk is None:
                break
            # import pdb;pdb.set_trace()
            response = chunk["choices"][0]["delta"]
            if 'content' not in response:
                continue
        
            full_response += response['content']
            full_response = full_response.replace('</s>', '')
        return full_response

        # return completion['choices'][0]['message']['content']
    if use_stream:
        result = show_completion(completion)
    else:
        result = completion.choices[0].message.content
    # print("ask_llm_anything result:", result)
    # tmp
    # reasoning = completion.choices[0].message.get("reasoning_content", "")
    # result = "<think>" + reasoning + "</think>" + "\n" + result
    if return_origin_response:
        return result, completion
    return result


def ask_llm_anything_with_retry(model_provider, model_name, messages,  args={
    "max_tokens": 256,
    "temperature": 0.5,
    "top_p": 1.0,
    "frequency_penalty": 0.0,
}, retry=3):
    min_sleep_time= 1
    max_sleep_time= 16
    current_sleep_time = min_sleep_time
    for i in range(retry):
        try:
            return ask_llm_anything(model_provider, model_name, messages, args)
        except Exception as e:
            print(f"Error: {e.args}")
            time.sleep(current_sleep_time)
            current_sleep_time = min(current_sleep_time * 2, max_sleep_time)
            continue
    return None

def ask_step_by_step(model_provider, model_name, messages, args={
    "max_tokens": 256,
    "temperature": 0.5,
    "top_p": 1.0,
    "frequency_penalty": 0.0,
}, retry=3):
    messages = deepcopy(messages)
    messages_to_ask = []
    
    msg_idx = 0 
    for msg in messages:
        if msg['role'].lower() in ["human", "user", "system"]:
            messages_to_ask.append(messages[msg_idx])
            msg_idx += 1
        else:
            answer = ask_llm_anything_with_retry(model_provider, model_name, messages_to_ask, args, retry)
            if answer is None:
                print("ask_step_by_step failed")
                messages_to_ask.append(messages[msg_idx])
            else:
                messages_to_ask.append({
                    "role": "assistant",
                    "content": answer
                })
            msg_idx += 1
    if messages_to_ask[-1]['role'] in ["human", "user"]:
        answer = ask_llm_anything_with_retry(model_provider, model_name, messages_to_ask, args, retry)
        if answer is not None:
            messages_to_ask.append({
                "role": "assistant",
                "content": answer
            })
        else:
            print("ask_step_by_step failed")
            messages_to_ask.append(
                {
                    "role": "assistant",
                    "content": None
                }
            )

    return messages_to_ask    



if __name__ == "__main__":

    # curl http://10.148.42.15:9200/v1/completions \
    # -H "Content-Type: application/json" \
    
    # -d '{
    #     "model": "step1u-summary-v0-int8",
    #     "prompt": "Say this is a test",
    #     "max_tokens": 7,
    #     "temperature": 0
    # }'

    # def complation_test(model_name, prompt, args = {}):
    #     import requests
    #     url = "http://stepcast-router.basemind-core.svc.platform.basemind.local:9200/v1/completions"
    #     headers = {
    #         "Content-Type": "application/json"
    #     }

    #     data = {
    #         "model": model_name,
    #         "prompt": prompt,
    #         "max_tokens": args.get("max_tokens", 100),
    #         "temperature": args.get("temperature", 0.5),
    #         "top_p": args.get("top_p", 1.0),
    #         "frequency_penalty": args.get("frequency_penalty", 0.0),
    #         # "stream": False
    #     }

    #     response = requests.post(url, headers=headers, data=json.dumps(data))

    #     print(response.json())

    # complation_test("step2p5t-it30k-sft0923-it1024", "Say this is a test", args={"max_tokens": 7, "temperature": 0})

    # pass
    messages = [
        {"role": "user", "content": "What is the capital of France?"}
    ]
    #answer = ask_llm_anything("openai", "doubao-1.5-ui-tars", messages, args={"max_tokens": 4096, "temperature": 0.2})
    answer = ask_llm_anything("openai", "claude-sonnet-4-5-20250929", messages, args={"max_tokens": 4096, "temperature": 0.2})
    print(answer)