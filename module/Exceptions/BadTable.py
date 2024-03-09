"""
Класс BadTable предоставляет возможность бросить исключение при неверно заполненной
таблице.

:param _message: Сообщение об ошибке.
Пример использования:

```python
>>> message = "fail"
>>> example = BadTable(message)
>>> example.what()
>>> print(example)

"""
class BadTable(Exception):
    
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
        return f"Ошибка в вводе данных таблицы: {self._message}"


    def what(self) -> str:
        return self._message