import os
from langchain_groq import ChatGroq
from langchain_openai import ChatOpenAI

def get_groq_llm(api_key: str = None, temperature: float = 0.3):
    """Returns a ChatGroq instance."""
    key = api_key or os.environ.get("GROQ_API_KEY")
    if not key:
        raise ValueError("Groq API Key is missing.")
    return ChatGroq(
        groq_api_key=key,
        model_name="llama-3.3-70b-versatile",
        temperature=temperature
    )

def get_openai_llm(api_key: str = None, temperature: float = 0.3):
    """Returns a ChatOpenAI instance."""
    key = api_key or os.environ.get("OPENAI_API_KEY")
    if not key:
        raise ValueError("OpenAI API Key is missing.")
    return ChatOpenAI(
        openai_api_key=key,
        model_name="gpt-4o-mini",
        temperature=temperature
    )
