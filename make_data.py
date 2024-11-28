import pandas as pd
import json


def read_json(file_name: str, file_path: str = "") -> None | list[dict] | dict:
    with open(f"{file_path}{file_name}.json", "r", encoding="utf-8") as r:
        return json.load(r)


def write_json(data: list[dict] | dict, file_name: str, file_path: str = "") -> None:
    with open(f"{file_path}{file_name}.json", "w", encoding="utf-8") as w:
        json.dump(data, w, indent=2, ensure_ascii=False)


def read_excel(file_name: str, file_path: str = "", file_extension: str = "xlsx") -> pd.DataFrame:
    return pd.read_excel(f"{file_path}{file_name}.{file_extension}")


def main():
    df = read_excel(file_name="Price_autodoc", file_path="files/")
    df.columns = ['art', 'maker', 'name', 'price', 'q_ty']
    new_catalog_info = {'KMK GLASS': {}, 'AGC': {}, 'BOR': {}, 'Nordglass': {}, 'БОР': {}}
    try:
        for item in df.itertuples():
            if item.maker == 'KMK GLASS':
                if item.art not in new_catalog_info['KMK GLASS']:
                    new_catalog_info['KMK GLASS'][item.art] = item.name
            elif item.maker == 'AGC':
                new_catalog_info['AGC'][item.art] = item.name
            elif item.maker == 'BOR':
                new_catalog_info['BOR'][item.art] = item.name
            elif item.maker == 'Nordglass':
                new_catalog_info['Nordglass'][item.art] = item.name
            elif item.maker == 'БОР':
                new_catalog_info['БОР'][item.art] = item.name
        write_json(data=new_catalog_info, file_name="info_codes", file_path="data/")
    except Exception as e:
        print(f"Ошибка при формировании файла с данными на этапе перебора позиций [38 строка]: {e}")


if __name__ == "__main__":
    main()
