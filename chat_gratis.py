import google.generativeai as genai

genai.configure(api_key="AIzaSyD6-HRefLzfS2EArUwyc6ux1CtgIYVsA0o")
model = genai.GenerativeModel('gemini-2.5-flash')

while True:
    user_input = input("Tú (Ingeniería): ")
    if user_input.lower() in ["salir", "exit"]: break
    
    response = model.generate_content(user_input)
    print(f"\nAntigravity AI: {response.text}\n")