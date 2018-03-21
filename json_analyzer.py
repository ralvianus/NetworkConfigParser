import json
from deepdiff import DeepDiff
import pprint

def get_json(file_name):
    with open(file_name) as json_file:
        json_data = json.load(json_file)
        return json_data

def compare_json(Hostname, Command, Data1, Data2):
    if (Data1 == Data2):
        print ("%s - %s output is same" % (Hostname, Command))
    else:
        print ("%s - %s output is different" % (Hostname, Command))
        pprint.pprint(DeepDiff(Data1, Data2))

def main():
    Hostname = raw_input('Input Hostname of the device : ').lower()
    Command = raw_input('Input Command : ').lower()
    Filename1 = raw_input('Input First JSON File : ').lower()
    Filename2 = raw_input('Input Second JSON File : ').lower()

    Data1 = get_json(Filename1)
    Data2 = get_json(Filename2)

    compare_json(Hostname, Command, Data1, Data2)

if __name__ == "__main__":
	# If this Python file runs by itself, run below command. If imported, this section is not run
	main()
