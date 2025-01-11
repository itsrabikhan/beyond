# IMPORTS

import math
import pandas as pd
from datetime import datetime


# CLASSES

class Colors:
    """
    This class contains ANSI escape codes for text colors and styles.
    """

    RED = "\u001b[31m"
    GREEN = "\u001b[32m"
    YELLOW = "\u001b[33m"
    BLUE = "\u001b[34m"
    MAGENTA = "\u001b[35m"
    CYAN = "\u001b[36m"
    WHITE = "\u001b[37m"
    GRAY = "\u001b[90m"
    RESET = "\u001b[0m"
    BOLD = "\u001b[1m"
    UNDERLINE = "\u001b[4m"

class Value:
    """
    This class represents a numerical value.
    Values exist for easy access to the value and its unit.
    """

    def __init__(self, value: int | float, unit: str):
        self.value = value
        self.unit = unit

    def __str__(self):
        return f"{self.value}{self.unit}"
    
    def __eq__(self, other):
        if type(other) == Value:
            return self.value == other.value and self.unit == other.unit
        if type(other) in [int, float]:
            return self.value == other
        return False
    
    def __ne__(self, other):
        return not self.__eq__(other)
    
    def __lt__(self, other):
        if type(other) == Value:
            if self.unit != other.unit:
                raise ValueError("Cannot compare values with different units, with the exception of equality or unitless values.")
            return self.value < other.value
        if type(other) in [int, float]:
            return self.value < other
        return False
    
    def __le__(self, other):
        return self.__lt__(other) or self.__eq__(other)
    
    def __gt__(self, other):
        if type(other) == Value:
            if self.unit != other.unit:
                raise ValueError("Cannot compare values with different units, with the exception of equality or unitless values.")
            return self.value > other.value
        if type(other) in [int, float]:
            return self.value > other
        return False
    
    def __ge__(self, other):
        return self.__gt__(other) or self.__eq__(other)
    
    def __add__(self, other):
        if type(other) == Value:
            if self.unit != other.unit:
                raise ValueError("Cannot add values with different units, with the exception of unitless values.")
            return Value(self.value + other.value, self.unit)
        if type(other) in [int, float]:
            return Value(self.value + other, self.unit)
        return None
    
    def __sub__(self, other):
        if type(other) == Value:
            if self.unit != other.unit:
                raise ValueError("Cannot subtract values with different units, with the exception of unitless values.")
            return Value(self.value - other.value, self.unit)
        if type(other) in [int, float]:
            return Value(self.value - other, self.unit)
        return None
    
    def __mul__(self, other):
        if type(other) == Value:
            if self.unit != other.unit:
                raise ValueError("Cannot multiply values with different units, with the exception of unitless values. This is a limitation of the current implementation.")
            return Value(self.value * other.value, self.unit)
        if type(other) in [int, float]:
            return Value(self.value * other, self.unit)
        return None
    
    def __truediv__(self, other):
        if type(other) == Value:
            if self.unit != other.unit:
                raise ValueError("Cannot divide values with different units, with the exception of unitless values. This is a limitation of the current implementation.")
            return Value(self.value / other.value, self.unit)
        if type(other) in [int, float]:
            return Value(self.value / other, self.unit)
        return None
    
    def __floordiv__(self, other):
        if type(other) == Value:
            if self.unit != other.unit:
                raise ValueError("Cannot divide values with different units, with the exception of unitless values. This is a limitation of the current implementation.")
            return Value(self.value // other.value, self.unit)
        if type(other) in [int, float]:
            return Value(self.value // other, self.unit)
        return None
    
    def __mod__(self, other):
        if type(other) == Value:
            if self.unit != other.unit:
                raise ValueError("Cannot divide values with different units, with the exception of unitless values. This is a limitation of the current implementation.")
            return Value(self.value % other.value, self.unit)
        if type(other) in [int, float]:
            return Value(self.value % other, self.unit)
        return None
    
    def __pow__(self, other):
        if type(other) == Value:
            if self.unit != other.unit:
                raise ValueError("Cannot raise values with different units, with the exception of unitless values. This is a limitation of the current implementation.")
            return Value(self.value ** other.value, self.unit)
        if type(other) in [int, float]:
            return Value(self.value ** other, self.unit)
        return None
    
    def __abs__(self):
        return Value(abs(self.value), self.unit)
    
    def __neg__(self):
        return Value(-self.value, self.unit)
    
    def __pos__(self):
        return Value(self.value, self.unit)
    
    def __int__(self):
        return int(self.value)
    
    def __float__(self):
        return float(self.value)
    
    def __round__(self, n: int = 0):
        return round(self.value, n)
    
    def __ceil__(self):
        return math.ceil(self.value)
    
    def __floor__(self):    
        return math.floor(self.value)

class Range:
    def __init__(self, start: Value, end: Value):
        if start.unit != end.unit:
            raise ValueError("Cannot create a range with different units.")
        if start > end:
            temp = start
            start = end
            end = temp
            del temp
        self.start = start
        self.end = end

class Operation:
    """
    This class represents an operation in the process.
    """

    def __init__(self, data: list):
        self.data = data

    def __str__(self) -> str:
        return str(self.data)
    
    def get_from(self) -> float:
        return float(self.data[0])
    
    def get_to(self) -> float:
        return float(self.data[1])
    
    def get_duration(self) -> float:
        return self.data[2]
    
    def get_operation(self) -> str:
        return self.data[4]
    
    def get_mud_weight(self) -> Value | None:
        operation = self.get_operation()
        try:
            data = operation.split("@")[1].split(". ")
            for entry in data:
                if "MW" in entry:
                    return Value(float(entry.replace("kg/m³ MW", "").strip()), "kg/m³")
            return None
        except:
            return None
        
class Drilling(Operation):
    """
    This class represents a drilling operation in the process.
    """

    def __str__(self) -> str:
        return str(f"[{convert_time(self.get_from())} - {convert_time(self.get_to())}] Drilling to {self.get_depth()} at {self.get_pump_rate()}. {self.get_mud_weight()} MW. {self.get_ecd()} ECD.")

    def get_depth(self) -> Value | None:
        try:
            operation = self.get_operation()
            return Value(float(operation.split("@")[0].replace("Drilling (to ", "").replace("m)", "")), "m")
        except:
            return None
    
    def get_ecd(self) -> Value | None:
        operation = self.get_operation()
        try:
            data = operation.split("@")[1].split(". ")
            for entry in data:
                if "ECD" in entry:
                    return Value(float(entry.replace("kg/m³ ECD", "").strip()), "kg/m³")
            return None
        except:
            return None
        
    def get_pump_rate(self) -> Value | None:
        operation = self.get_operation()
        try:
            data = operation.split("@")[1].split(". ")
            for entry in data:
                if "m³/min" in entry:
                    return Value(float(entry.replace("m³/min", "").strip()), "m³/min")
            return None
        except:
            return None

class Connection(Operation):
    """
    This class represents a connection operation in the process.
    """

    def __str__(self) -> str:
        return str(f"[{convert_time(self.get_from())} - {convert_time(self.get_to())}] Connection at {self.get_depth()}. {self.get_mud_weight()} MW. {self.get_esd()} ESD.")

    def get_depth(self) -> Value | None:
        try:
            operation = self.get_operation()
            return Value(float(operation.split("@")[1].split(". ")[0].replace("m", "").strip()), "m")
        except:
            return None
    
    def get_esd(self) -> Value | None:
        try:
            operation = self.get_operation()
            data = operation.split("@")[1].split(". ")
            for entry in data:
                if "ESD" in entry:
                    return Value(float(entry.replace("kg/m³ ESD", "").strip()), "kg/m³")
            return None
        except:
            return None
        
    def get_gas(self) -> Value | None:
        try:
            operation = self.get_operation()
            data = operation.split("@")[1].split(". ")
            for entry in data:
                if "KSCM/Day B/U" in entry:
                    return Value(float(entry.replace("KSCM/Day", "").strip()), "KSCM/Day")
                elif "No Gas to Report" in entry:
                    return Value(0, "KSCM/Day")
            return None
        except:
            return None
        
    def get_static_bp(self) -> Value | Range | None:
        try:
            operation = self.get_operation()
            data = operation.split("@")[1].split(". ")
            for entry in data:
                if "BP" in entry:
                    if any(i.lower() in entry.lower() for i in ["No BP", "Closed Choke", "Open Choke"]):
                        return Value(0, "kPa")
                    elif " to " in entry:
                        "1400 to 1650kPa BP"
                        values = entry.replace("kPa BP", "").split(" to ")
                        start = Value(float(values[0]), "kPA")
                        end = Value(float(values[1]), "kPA")
                        return Range(start, end)
            return None
        except:
            return None


# UTILITY FUNCTIONS

def fprint(*args, end: str = "\n") -> None:
    """
    Print function with built-in color reset for convenience.
    This is to be used instead of the built-in print function because of its automatic color reset.

    Parameters:
        *args (any): Any number of arguments to print, can be any type.

    Returns:
        None
    """

    # Altered print statement to include color reset.
    print(str(*args) + Colors.RESET, end=end)

def clear() -> None:
    """
    Clears the terminal screen using an ANSI escape code.

    Parameters:
        None
    
    Returns:
        None
    """

    # ANSI escape code to clear the screen.
    print("\033[H\033[J", end='')

def pause(newlines: int = 1) -> None:
    """
    Pause the program until the user presses Enter.

    Parameters:
        newlines (int): The number of newlines to print before the prompt. Default is 1.

    Returns:
        None
    """

    # Print newlines and prompt the user to press the enter key.
    fprint(("\n" * newlines) + "Press ENTER to continue.", end="")
    input()

def convert_time(time_object: str | float) -> str | float:
    """
    Converts time from float to string or vice versa.
    Time can be represented as a float between 0 to 1 or a string in the format "HH:MM".

    Parameters:
        time_object (str | float): The time in either format.

    Returns:
        str | float: The time in the other format.
    """

    if type(time_object) == str:
        hour = int(time_object.split(":")[0])
        minute = int(time_object.split(":")[1])
        return float((hour * 60 + minute) / 1440)
    elif type(time_object) == float:
        hour = int(time_object * 1440 / 60)
        minute = int(time_object * 1440 % 60)
        return f"{hour}:{minute:02d}"
    else:
        return None


# MAIN FUNCTION

def main() -> None:
    data = pd.ExcelFile('sheet.xlsb')
    sheet_names = data.sheet_names
    sheets = []
    for i in sheet_names:
        try:
            datetime.strptime(i, '%Y %b-%d')
            sheets.append(i)
        except ValueError:
            continue
        
    clear()
    print(sheets[2])
    parsed = data.parse(sheets[2], dtype=str)
    data_list = parsed.values.tolist()
    new = []
    count = 0
    for i in data_list:
        count += 1
        if count <= 21:
            continue
        sub = []
        for j in i:
            if type(j) != float:
                sub.append(j)
                #print(type(j), end=' | ')
        new.append(sub)
        #print()

    if len(new) < 0:
        fprint(f"{Colors.RED}ERROR CODE 0{Colors.RESET}")
        return
    
    operations = []
    for index, row in enumerate(new):
        if index == 0 and row[0] != "0":
            fprint(f"{Colors.RED}ERROR CODE 1{Colors.RESET}")
            return
        
        if len(row) <= 1:
            break
        
        try:
            cnt = int(row[3])
            if cnt > 0:
                operations.append(Connection(row))
                continue
        except ValueError:
            pass

        if row[3] == "C":
            operations.append(Connection(row))
        elif row[3] == "D":
            operations.append(Drilling(row))

    for i in operations:
        print(str(i))
    

if __name__ == "__main__":
    main()