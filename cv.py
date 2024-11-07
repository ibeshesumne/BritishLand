def get_input(section_name, prompt, multiple_entries=False):
    print(f"\n--- {section_name} ---")
    entries = []
    while True:
        entry = input(prompt)
        if entry:
            entries.append(entry)
            if not multiple_entries:
                break
        else:
            break
    return entries

def main():
    # Collecting information
    personal_statement = get_input("Personal Statement", "Enter your personal statement: ")[0]
    key_skills = get_input("Key Skills", "Enter a skill (leave blank to end): ", multiple_entries=True)
    employment_history = get_input("Employment History", "Enter an employment detail (leave blank to end): ", multiple_entries=True)
    education = get_input("Education", "Enter an education detail (leave blank to end): ", multiple_entries=True)
    hobbies_interests = get_input("Hobbies and Interests", "Enter a hobby or interest (leave blank to end): ", multiple_entries=True)

    # Writing to a file
    with open("cv.txt", "w") as file:
        file.write("Curriculum Vitae\n\n")
        file.write("Personal Statement:\n")
        file.write(personal_statement + "\n\n")
        
        file.write("Key Skills:\n")
        file.write("\n".join(key_skills) + "\n\n")

        file.write("Employment History:\n")
        file.write("\n".join(employment_history) + "\n\n")

        file.write("Education:\n")
        file.write("\n".join(education) + "\n\n")

        file.write("Hobbies and Interests:\n")
        file.write("\n".join(hobbies_interests) + "\n")

    print("\nCV created and saved as cv.txt.")

if __name__ == "__main__":
    main()

