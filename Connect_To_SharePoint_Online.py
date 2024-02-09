import io, pandas as pd, tempfile
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.folders.folder import Folder
from office365.sharepoint.files.file import File


# SharePoint Credentials
sharepoint_clientId = "XXXXXXXXXXXXXXXXXX"
sharepoint_clientsecret = "XXXXXXXXXXXXXXXXXXXXXXXX"


class SharePoint_Connection:
    def __init__(self, client_id, client_secret, team):
        self.client_id = client_id
        self.client_secret = client_secret
        self.team = team

    def establish_sharepoint_context(self):
        try:
            site_url = f"https://YOUR_COMPANY.sharepoint.com/teams/{self.team}"
            context_auth = AuthenticationContext(site_url)
            if context_auth.acquire_token_for_app(
                client_id=self.client_id, client_secret=self.client_secret
            ):
                ctx = ClientContext(site_url, context_auth)
                return ctx
        except Exception as e:
            print(
                f"error executing read_sharepoint_file_as_df(): {type(e).__name__} {e}"
            )

    def create_sharepoint_directory(self, directory_name: str):
        """
        Creates a folder in the sharepoint directory.
        """
        if directory_name:
            ctx = self.establish_sharepoint_context()
            result = ctx.web.folders.add(
                f"Shared Documents/General/{directory_name}"
            ).execute_query()
            if result:
                # documents is titled as Shared Documents for relative URL in SP
                relative_url = f"Shared Documents/General/{directory_name}"
                print(relative_url)
                return relative_url
            else:
                print("folder could not be created")

    def read_sharepoint_file_as_df(self, file_path, dtype=None):
        ctx = self.establish_sharepoint_context()
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()

        out = io.BytesIO()
        f = (
            ctx.web.get_file_by_server_relative_url(f"/Shared Documents/{file_path}")
            .download(out)
            .execute_query()
        )
        if dtype is not None:
            # df = pd.read_csv(out, dtype=dtype)
            df = pd.read_excel(out, dtype=dtype)
        else:
            # df = pd.read_csv(out)
            df = pd.read_excel(out)
        out.close()
        return df

    def write_bytefile_to_sharepoint(
        self, file_path: str, file_name: str, file_bytes: bytes
    ) -> None:
        ctx = self.establish_sharepoint_context()

        folder: Folder = ctx.web.get_folder_by_server_relative_url(
            f"Shared Documents/{file_path}"
        )

        chunk_size: int = 5000

        # Check if the file already exists
        file: File = folder.files.get_by_url(file_name)

        if file.exists:

            # If the file exists, delete it
            file.delete_object().execute_query()

        # Create a temporary file and write the bytes to it
        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
            temp_file.write(file_bytes)

        # Use the temporary file for uploading
        with open(temp_file.name, "rb") as file_to_upload:
            folder.files.create_upload_session(
                file=file_to_upload, chunk_size=chunk_size, file_name=file_name
            ).execute_query()
        print(f"File has been uploaded successfully!")


connectionString = SharePoint_Connection(
    client_id=sharepoint_clientId,
    client_secret=sharepoint_clientsecret,
    team="[Your SharePoint Team's Name]",
)


# create directory
# connectionString.create_sharepoint_directory("TestFolder")


# # read sharepoint excel file
# df = connectionString.read_sharepoint_file_as_df(
#     "General/folder/file_name.xlsx",
# )
# print(df.head())

# create pandas dataframe to test upload
df = pd.DataFrame(data={"col1": [1, 2], "col2": [3, 4], "col3": [5, 6]})
fileBytes = bytes(df.to_csv(index=False), encoding="utf-8")


connectionString.write_bytefile_to_sharepoint(
    "General/folder_root/sub_folder", "test2.csv", fileBytes
)
