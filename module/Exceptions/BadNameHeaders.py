"""
Класс BadNameHeaders предоставляет возможность бросить исключение при неверно заполненной
таблице.

:param _message: Сообщение об ошибке.
Пример использования:

```python
>>> message = "fail"
>>> example = BadNameHeaders(message)
>>> example.what()
>>> print(example)

"""
class BadNameHeaders(Exception):
    
    """
    Конструктор создания исключения.

    Args:
        message: str Сообщение, содержащее более подробную информацию об
        исключении;
    Returns:
        None;
    """
    def __init__(self, message: str):
        self._message = message;
    
    def __str__(self):
        return f"Неправильные имена заголовков: {self._message}"


    def what(self) -> str:
        return self._message