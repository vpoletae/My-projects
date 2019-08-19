from transliterate import translit
from googletrans import Translator

def translate_to_rus(word):
      translator = Translator()
      translate = translator.translate(word, src='en', dest='ru').text
      
      transliteartion = translit(word, language_code = "ru")

      return (transliteartion, translate)

def translate_to_eng(word):
      translator = Translator()
      translate = translator.translate(word, src='ru', dest='en').text
      
      transliteartion = translit(word, language_code = "ru", reversed =True)

      return (transliteartion, translate)

