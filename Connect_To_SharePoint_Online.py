import io, pandas as pd, tempfile
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.folders.folder import Folder
from office365.sharepoint.files.file import File


# SharePoint Credentials
sharepoint_clientId = "XXXXXXXXXXXXXXXXXX"
sharepoint_clientsecret = "XXXXXXXXXXXXXXXXXXXXXXXX"


class SharePoint_Connection:
    def __init__(self, client_id: str, client_secret: str, team: str) -> None:
        """
        Constructor to initialize SharePoint_Connection object with client_id, client_secret, and team values.

        Parameters:
            - client_id (str): The client ID used for SharePoint authentication.
            - client_secret (str): The client secret used for SharePoint authentication.
            - team (str): The name of the SharePoint team.

        Returns:
            - None
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.team = team

    def establish_sharepoint_context(self):
        """
        Establishes a SharePoint context using the provided client_id, client_secret, and team.

        Parameters:
            - self: An instance of the class that contains the method.

        Returns:
            - ctx (ClientContext): A SharePoint client context established using the provided authentication credentials
            (client_id, client_secret) and team information. If successful, the function returns the ClientContext object.
            If an error occurs during the establishment of the SharePoint context, an exception is caught and an error message
            is printed, and the function returns None.
        """
        # Establishes a SharePoint context using the provided client_id, client_secret, and team.
        try:
            site_url = f"https://YOUR_COMPANY.sharepoint.com/teams/{self.team}"
            context_auth = AuthenticationContext(site_url)
            if context_auth.acquire_token_for_app(
                client_id=self.client_id, client_secret=self.client_secret
            ):
                ctx = ClientContext(site_url, context_auth)
                return ctx
        except Exception as e:
            print(f"Error: {type(e).__name__} {e}")

    def create_sharepoint_directory(self, directory_name: str) -> None:
        """
        Creates a directory in SharePoint under the 'Shared Documents/General/' path.

        Parameters:
            - directory_name (str): The name of the directory to be created.

        Returns:
            - str: The relative URL of the created directory if successful. If an error occurs during the creation process,
            an error message is printed, and the function returns an empty string.
        """
        if directory_name:
            # Establish SharePoint context
            ctx = self.establish_sharepoint_context()
            # Attempt to create the directory

            result = ctx.web.folders.add(
                f"Shared Documents/General/{directory_name}"
            ).execute_query()
            # If successful, return the relative URL of the created directory
            if result:
                relative_url = f"Shared Documents/General/{directory_name}"
                print(relative_url)
                return relative_url
            else:
                print("Fail to create a folder/directory!")
                return ""

    def read_sharepoint_file_as_df(self, file_path: str, dtype=None) -> pd.DataFrame:
        """
        Reads a file from SharePoint and returns its content as a Pandas DataFrame.

        Parameters:
            - file_path (str): The path of the file in SharePoint, relative to the 'Shared Documents' directory.
            - dtype (dict or None): Data type specification for columns in the DataFrame (optional).

        Returns:
            - pd.DataFrame: A Pandas DataFrame containing the content of the specified file. If the 'dtype' parameter is provided,
            it is used to specify data types for DataFrame columns during reading.
        """
        # # Establish SharePoint context
        ctx = self.establish_sharepoint_context()
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        # Download file content
        out = io.BytesIO()
        f = (
            ctx.web.get_file_by_server_relative_url(f"/Shared Documents/{file_path}")
            .download(out)
            .execute_query()
        )
        # Read file content into Pandas DataFrame
        if dtype is not None:
            # # If data types are specified, use them during DataFrame creation
            # df = pd.read_csv(out, dtype=dtype)
            df = pd.read_excel(out, dtype=dtype)
        else:
            # Otherwise, read the file without specifying data types
            # df = pd.read_csv(out)
            df = pd.read_excel(out)
        # Close the BytesIO stream
        out.close()
        return df

    def write_bytefile_to_sharepoint(
        self,
        file_path: str,
        file_name: str,
        file_bytes: bytes,
    ) -> None:
        """
        Writes a byte file to SharePoint in the specified folder with the given file name.

        Parameters:
            - file_path (str): The path of the folder in SharePoint where the file should be written, relative to the 'Shared Documents' directory.
            - file_name (str): The name to be given to the file in SharePoint.
            - file_bytes (bytes): The content of the file as a bytes object.

        Returns:
            - None
        """
        # # Establish SharePoint context
        ctx = self.establish_sharepoint_context()
        # Get the SharePoint folder by server-relative URL
        folder: Folder = ctx.web.get_folder_by_server_relative_url(
            f"Shared Documents/{file_path}"
        )
        # # Chunk size for uploading
        chunk_size: int = 500000

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
        # Print success message
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

# create pandas data frame to test upload
df = pd.DataFrame(data={"col1": [1, 2], "col2": [3, 4], "col3": [5, 6]})
fileBytes = bytes(df.to_csv(index=False), encoding="utf-8")


connectionString.write_bytefile_to_sharepoint(
    "General/folder_root/sub_folder", "test2.csv", fileBytes
)
