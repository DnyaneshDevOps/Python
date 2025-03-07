def person_info(**details):
    for key, value in details.items():
        print(f"{key}: {value}")

person_info(name="Alice", age=25, city="New York")