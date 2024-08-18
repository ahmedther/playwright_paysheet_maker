import pandas as pd
from .helpers import Helpers


class PandasHandler(Helpers):
    def __init__(self):
        super().__init__()

    def convert_to_dataframe(self, data):
        df = pd.DataFrame(data, columns=["Reference", "Date of Session"])
        df["Client (Full Name)"] = df["Reference"].apply(self.extract_client_name)
        df["Rate"] = 220
        df["Total Hours"] = df["Reference"].apply(
            lambda x: (
                1
                if "60 min" in x or "60 mins" in x
                else 1.5 if "90 min" in x or "90 mins" in x else 0
            )
        )
        # Reordering columns
        df = df[["Reference", "Client (Full Name)", "Rate", "Date of Session"]]

        # Adding three more empty columns named "Date of Session"
        df["Date of Session 2"] = ""
        df["Date of Session 3"] = ""
        df["Date of Session 4"] = ""

        # Renaming the columns to have the same name
        df.columns = [
            "Reference",
            "Client (Full Name)",
            "Rate",
            "Date of Session",
            "Date of Session",
            "Date of Session",
            "Date of Session",
        ]
        df["Total Hours"] = df["Reference"].apply(
            lambda x: (
                1
                if "60 min" in x or "60 mins" in x
                else 1.5 if "90 min" in x or "90 mins" in x else 0
            )
        )

        df["Amount To Be Paid"] = ""
        df["For Office Use Only"] = ""

        return df

    def concat_dataframes(self, args):
        return pd.concat(args, ignore_index=True)
