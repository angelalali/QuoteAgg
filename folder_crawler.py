import os

def list_files(dir):
    r = []
    for root, dirs, files in os.walk(dir):
        for name in files:
            r.append(os.path.join(root, name))
    return r

# path = '//teslamotors.com/US/Finance/Cost IQ/Quotes'
## must start the path with a backward slash!!! otherwise it's not a valid path and nothing will be read

# for path in list_files(path):
#     print(path)


# import os
# rootdir = 'C:/Users/yisli/Documents/landlordlady/Tesla/projects/'
#
# for subdir, dirs, files in os.walk(rootdir):
#     for file in files:
#         print(os.path.join(subdir, file))

