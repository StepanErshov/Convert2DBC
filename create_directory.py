import os
import getpass


class DirectoryCreator:
    """Класс для создания иерархии директорий согласно заданной структуре."""

    def __init__(self):
        self.HIERARCHI = {
            "04.01.01.Body CAN": [
                "04.01.01.01_BCM",
                "04.01.01.02_DCM",
                "04.01.01.03_HLL",
                "04.01.01.04_HLR",
                "04.01.01.05_PAT",
                "04.01.01.06_SCU",
                "04.01.01.08_WCBS",
            ],
            "04.01.02.Powertrain CAN": [
                "04.01.02.01_APTC",
                "04.01.02.02_BMS",
                "04.01.02.03_EVCOM",
                "04.01.02.04_IBS_PT",
                "04.01.02.05_MCU",
                "04.01.02.06_POD",
                "04.01.02.07_PRND",
                "04.01.02.08_TCU",
                "04.01.02.09_VCU_PT",
            ],
            "04.01.04.Entertainment CANFD": [
                "04.01.04.01_DIM",
                "04.01.04.02_HOD_Heating",
                "04.01.04.03_HUD",
                "04.01.04.04_IVI_SFI",
                "04.01.04.05_MFP",
                "04.01.04.06_Switches",
                "04.01.04.07_SWP",
            ],
            "04.01.05.Chassis CANFD": [
                "04.01.05.01_ACU",
                "04.01.05.02_AVAS",
                "04.01.05.03_EPS",
                "04.01.05.04_IBS_CH",
                "04.01.05.05_VCU_CH",
            ],
            "04.01.06.Demilitary zone CANFD": [
                "04.01.04.01_NDT",
                "04.01.06.02_CBM",
                "04.01.06.03_ERA",
            ],
            "04.01.07.CGW,SGW,ADCU": [
                "04.01.07.01_ADCU",
                "04.01.07.02_CGW",
                "04.01.07.03_SGW",
            ],
            "04.01.08 Diagnostic CAN": [
                "04.01.08.01_DTOOL",
            ]
        }
        self.USERNAME = getpass.getuser()
        self.PATH_DOC = f"C:\\Users\\{self.USERNAME}\\Documents"

    def create_directory_structure(self):
        for dir, pod_dir in self.HIERARCHI.items():
            os.makedirs(f"{self.PATH_DOC}\\{dir}", exist_ok=True)
            for small_dir in pod_dir:
                os.makedirs(f"{self.PATH_DOC}\\{dir}\\{small_dir}", exist_ok=True)
        print("Директории успешно созданы!")

    def get_hierarchy(self):
        return self.HIERARCHI

    def set_hierarchy(self, new_hierarchy: dict):
        self.HIERARCHI = new_hierarchy

    def set_custom_path(self, custom_path: str):
        self.PATH_DOC = custom_path


# if __name__ == "__main__":
creator = DirectoryCreator()
creator.create_directory_structure()

    # creator.set_custom_path("D:\\MyCustomPath")
    # creator.set_hierarchy({"NewParent": ["Child1", "Child2"]})
    # creator.create_directory_structure()
