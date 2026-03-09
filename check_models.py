import google.generativeai as genai

GEMINI_API_KEY = "AIzaSyCRKMvEp_afVZBdzdtDrGjEMyA1pnS2fW4"
genai.configure(api_key=GEMINI_API_KEY)

for model in genai.list_models():
    if "generateContent" in model.supported_generation_methods:
        print(model.name)


