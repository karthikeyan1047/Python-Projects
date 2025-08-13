import inspect
import importlib


# # LIST FUNCTIONS IN A MODULE
# # ``````````````````````````````
# import selenium.webdriver.common     # change the module to inspect
# module = selenium.webdriver.common    # change the module to inspect

# functions_list = inspect.getmembers(module, lambda f: inspect.isfunction(f) or inspect.isbuiltin(f))

# print(len(functions_list))
# print()

# for function, _ in functions_list:
#     print(function)

# # METHODS AND PROPERTIES
# ``````````````````````````
package_name = "selenium.webdriver.common"

try:
    package = importlib.import_module(package_name)
    items = dir(package)
    for item in items:
        print(item)

except ModuleNotFoundError:
    print(f"Package '{package_name}' not found. Make sure it is installed.")
