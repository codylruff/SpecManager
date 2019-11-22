import os

def main():
    new_rba_dir = r"S:\Production\Protection Production\Setup Sheets\RBA's\New RBAs"

    for filename in os.listdir(new_rba_dir):
            print(os.path.join(new_rba_dir, filename))

if __name__ == "__main__":
    main()
