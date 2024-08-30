from annotated_types import LowerCase
from openai import OpenAI 
from CargarJson import cargar_configuracion

archivo_config = './config.json'
configuracion = cargar_configuracion(archivo_config)
key=configuracion['key']

def GetGPTResponse(texto):
 openai= OpenAI(api_key=key) 
 formula = texto[0]
 formula = formula.lower()
 texto=list(texto)
 if texto[2] == '' and len(texto[1]) == 1:
  texto[2] = texto[1]
 if texto[0] == 'SUMA ENTRE DOS CELDAS' or texto[0] == 'SUMA':
   prompt = formula + ' de la celda ' + texto[1] + ' y la celda ' + texto[2] + ' no incluyas "SUMA" o "SUM" solo usa el signo "=" y "+" '
  
 else:
   prompt = formula + ' desde la celda ' + texto[1] + ' hasta la celda ' + texto[2]
  
 print(prompt)

 if texto[0] == '' or texto[1] == '' or texto[2] == '' :
   respuesta='nan'
   
 if texto[0] == 'SUMA POR RANGOS' or texto[0] == 'MULTIPLICACION' or texto[0] == 'PROMEDIO':
   response = openai.chat.completions.create(
     model='gpt-3.5-turbo', 
     messages=[
     {"role": "system", "content": "Dame solo la formula de excel en formato ingles y reemplaza el ';' por ':' en las formulas, no añadas ninguna palabra mas a tu respuesta."},
     {"role": "user", "content": prompt}
   ],     
     temperature=0.8,         
     max_tokens=50              
   )
 elif texto[0] == 'CONTAR':
    response = openai.chat.completions.create(
     model='gpt-3.5-turbo', 
     messages=[
     {"role": "system", "content": "Dame solo la formula de excel en formato ingles (COUNT) y reemplaza el ';' por ':', No olvides la estructura de una formula de excel en las formulas, no añadas ninguna palabra mas a tu respuesta."},
     {"role": "user", "content": prompt}
   ],     
     temperature=0.8,         
     max_tokens=50              
   )
 else:
   response = openai.chat.completions.create(
     model='gpt-3.5-turbo', 
     messages=[
     {"role": "system", "content": "Dame solo la formula de excel, no añadas ninguna palabra mas a tu respuesta."},
     {"role": "user", "content": prompt}
   ],     
     temperature=0.8,         
     max_tokens=50              
   )

 print(response.choices[0].message.content)
 respuesta=response.choices[0].message.content

 return respuesta, texto
