from os import makedirs as os_makedirs, path as os_path
import calendar

def makeFileSystemByYear_to_output(year):

    #make output_src year folder, and the .gitkeep file
    os_makedirs( os_path.join("output_src", "____{}".format(str(year)[2:])), exist_ok=True )
    with open(os_path.join("output_src", "____{}".format(str(year)[2:]), ".gitkeep"), 'w') as f:
        pass


    #make output_src year month folder
    for month in range(1,13):
        os_makedirs( os_path.join("output_src", "____{}".format(str(year)[2:]), "__{}{}".format(str(month).zfill(2), str(year)[2:])), exist_ok=True )
        with open(os_path.join("output_src", "____{}".format(str(year)[2:]), "__{}{}".format(str(month).zfill(2), str(year)[2:]), ".gitkeep"), 'w') as f:
            pass

    #make output_src year month day folder
    for month in range(1,13):
        for day in range(1, calendar.monthrange(year, month)[1]+1):
            os_makedirs( os_path.join("output_src", "____{}".format(str(year)[2:]), "__{}{}".format(str(month).zfill(2), str(year)[2:]), "{}{}{}".format(str(day).zfill(2), str(month).zfill(2), str(year)[2:])), exist_ok=True )
            with open(os_path.join("output_src", "____{}".format(str(year)[2:]), "__{}{}".format(str(month).zfill(2), str(year)[2:]), "{}{}{}".format(str(day).zfill(2), str(month).zfill(2), str(year)[2:]), ".gitkeep"), 'w') as f:
                pass

            os_makedirs( os_path.join("output_src", "____{}".format(str(year)[2:]), "__{}{}".format(str(month).zfill(2), str(year)[2:]), "{}{}{}".format(str(day).zfill(2), str(month).zfill(2), str(year)[2:]), "byEnterprise"), exist_ok=True )
            with open(os_path.join("output_src", "____{}".format(str(year)[2:]), "__{}{}".format(str(month).zfill(2), str(year)[2:]), "{}{}{}".format(str(day).zfill(2), str(month).zfill(2), str(year)[2:]), "byEnterprise", ".gitkeep"), 'w') as f:
                pass

            os_makedirs( os_path.join("output_src", "____{}".format(str(year)[2:]), "__{}{}".format(str(month).zfill(2), str(year)[2:]), "{}{}{}".format(str(day).zfill(2), str(month).zfill(2), str(year)[2:]), "byExamCategory"), exist_ok=True )
            with open(os_path.join("output_src", "____{}".format(str(year)[2:]), "__{}{}".format(str(month).zfill(2), str(year)[2:]), "{}{}{}".format(str(day).zfill(2), str(month).zfill(2), str(year)[2:]), "byExamCategory", ".gitkeep"), 'w') as f:
                pass


def makeFileSystemByYear_to_input(year):

    #make output_src year folder, and the .gitkeep file
    os_makedirs( os_path.join("input_src", "____{}".format(str(year)[2:])), exist_ok=True )
    with open(os_path.join("input_src", "____{}".format(str(year)[2:]), ".gitkeep"), 'w') as f:
        pass


    #make input_src year month folder
    for month in range(1,13):
        os_makedirs( os_path.join("input_src", "____{}".format(str(year)[2:]), "__{}{}".format(str(month).zfill(2), str(year)[2:])), exist_ok=True )
        with open(os_path.join("input_src", "____{}".format(str(year)[2:]), "__{}{}".format(str(month).zfill(2), str(year)[2:]), ".gitkeep"), 'w') as f:
            pass

if __name__ == "__main__":
    makeFileSystemByYear_to_output(2026)
    #makeFileSystemByYear_to_input(2026)