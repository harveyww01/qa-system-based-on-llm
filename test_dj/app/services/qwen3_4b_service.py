"""
Qwen3-4B服务：调用本地部署的Qwen3-4B模型进行文本生成。
"""
from openai import OpenAI
from openai import APIConnectionError, APIError


def local_chat(message, model_config, model_name=None):
    """
    调用本地Qwen3-4B模型生成文本。
    
    Args:
        message: 用户输入消息
        model_config: 模型配置参数（temperature, max_tokens等）
        model_name: 模型名称
    
    Returns:
        str: 模型生成的响应文本
    """
    base_url = "https://u964814-962f-1298fd9f.bjb1.seetacloud.com:8443/v1"
    model = 'qwen3-4b-lora-fp8'

    client = OpenAI(base_url=base_url, api_key="no-key")

    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "你是一个专业的中文助手，请始终使用简体中文回答所有问题，不要使用英文。"},
                {"role": "user", "content": message}
            ],
            temperature=model_config.get("temperature", 0.2),
            max_tokens=model_config.get("max_tokens", 8192),
            top_p=model_config.get("top_p", 0.9),
            frequency_penalty=model_config.get("frequency_penalty", 0.1),
            stream=False
        )
        return resp.choices[0].message.content.strip()

    except APIConnectionError:
        return "错误：无法连接到模型服务！"
    except APIError as e:
        return f"错误：模型调用失败 - {str(e)}"
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print(f"完整错误堆栈:\n{error_trace}")
        return f"错误：服务器错误，生成失败 - {str(e)}"