import io, pandas as pd, tempfile
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.folders.folder import Folder
from office365.sharepoint.files.file import File


# SharePoint Credentials
sharepoint_clientId = "XXXXXXXXXXXXXXXX"
sharepoint_clientsecret = "XXXXXXXXXXXXXXXXXXXXXXX"


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

        try:
            # SharePoint site URL based on the company's domain name and team
            site_url = f"https://<example>.sharepoint.com/teams/{self.team}"
            # Authentication context using client_id and client_secret
            context_auth = AuthenticationContext(site_url)
            # Acquire token for the application
            if context_auth.acquire_token_for_app(
                client_id=self.client_id, client_secret=self.client_secret
            ):
                # Create SharePoint client context
                ctx = ClientContext(site_url, context_auth)
                return ctx
        except Exception as e:
            # Print error message if an exception occurs during SharePoint context establishment
            print(f"Error: {type(e).__name__} {e}")
            return None

    def create_sharepoint_directory(self, directory_name: str) -> str | None:
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
            try:
                result = ctx.web.folders.add(
                    f"Shared Documents/General/{directory_name}"
                ).execute_query()
                # If successful, return the relative URL of the created directory
                if result:

                    relative_url = f"Shared Documents/General/{directory_name}"
                    print(
                        f"{directory_name} directory has been created at '{relative_url}'"
                    )
                    return relative_url
                else:
                    print("Failed to create a folder/directory!")
                    return ""
            except Exception as e:
                # Print error message if an exception occurs during directory creation
                print(f"Error: {type(e).__name__} {e}")
                return ""
        else:
            print("Directory name cannot be empty!")
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
        # Establish SharePoint context
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
            # If data types are specified, use them during DataFrame creation
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
        self, file_path: str, file_name: str, file_bytes: bytes
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
        # Establish SharePoint context
        ctx = self.establish_sharepoint_context()
        # Get the SharePoint folder by server-relative URL
        folder: Folder = ctx.web.get_folder_by_server_relative_url(
            f"Shared Documents/{file_path}"
        )
        # Chunk size for uploading
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
        print(f"{file_name} has been uploaded successfully!")


# Establish Connection
connectionString = SharePoint_Connection(
    client_id=sharepoint_clientId,
    client_secret=sharepoint_clientsecret,
    team="Your Company's Team Name",
)

########################################################################
# Create directory
connectionString.create_sharepoint_directory("folder_name")

#########################################################################

# Read sharepoint files
df = connectionString.read_sharepoint_file_as_df(
    "General/folder/file_name.xlsx",
)
print(df.head())

########################################################################
# Write files to SharePoint Online
df = pd.DataFrame(data={"col1": [1, 2], "col2": [3, 4], "col3": [5, 6]})
fileBytes = bytes(df.to_csv(index=False), encoding="utf-8")


connectionString.write_bytefile_to_sharepoint(
    "General/folder", "file_name.csv", fileBytes
)
