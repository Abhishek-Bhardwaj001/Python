from typing import Any, Dict, List, Optional, Union
import pandas as pd
import io

# Library imports
from src.utils.library_import import (
    requests,
    ConfidentialClientApplication,
    time,
    logging,
    urlparse,
    unquote,
    parse_qs,
    json
)

# Exception imports
from src.exceptions import (
    AuthenticationError,
    DownloadError,
    APIError
)

class SharepointConnector():
    def __init__(
        self, 
        tenant_id: str, 
        client_id: str, 
        client_secret: str, 
        authority: str, 
        graph: str, 
        scope: List[str],
        tenant: str
    ) -> None:
        """
        Initialize SharepointConnector with authentication and configuration parameters.

        Args:
            tenant_id (str): Azure tenant ID.
            client_id (str): Azure client ID.
            client_secret (str): Azure client secret.
            authority (str): Azure authority URL.
            graph (str): Microsoft Graph API endpoint.
            scope (List[str]): List of scopes for authentication.
            tenant (str): SharePoint tenant name.
        """
        self.headers = None
        self.token_last_fetch = 0
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.authority = authority
        self.graph = graph
        self.scope = scope
        self.tenant=tenant
        self._get_headers()
    
# ============================================: List Item Methods :===============================================

    def get_site_drives(self, site_name: str, verbose: bool = False) -> pd.DataFrame:
        """
        List all drives for a given SharePoint site.

        Args:
            site_name (str): Name of the SharePoint site.
            verbose (bool, optional): If True, print drive details. Defaults to False.

        Returns:
            pd.DataFrame: DataFrame containing drive metadata.
        """
        try:
            site_id = self._get_site_id(f"{self.tenant}/sites/{site_name}/",verbose=verbose)
            resp = requests.get(f"{self.graph}/sites/{site_id}/drives", headers=self.headers)
            resp = resp.json()
            all_drives = []
            drives = resp["value"]
            for d in resp["value"]:
                if verbose:
                    print(f"Drive Name: {d.get('name',None)} | ID: {d.get('id',None)} | Drive Type: {d.get('driveType',None)} | Created At: {d.get('createdDateTime',None)} | Created By: {d.get('createdBy',None).get('user',None)} | Modified At: {d.get('lastModifiedDateTime',None)} | Web Url: {d.get('webUrl',None)}")
                all_drives.append({
                    "id": d.get('id',None),
                    "drive_name": d.get('name',None),
                    "weburl": d.get('webUrl',None),
                    "drive_Type": d.get('driveType',None),
                    "created_at": d.get('createdDateTime',None),
                    "created_by": d.get('createdBy',None).get('user',None),
                    "modified_at": d.get('lastModifiedDateTime',None),
                    "site_name": site_name
                })
            return pd.DataFrame(all_drives)
        except Exception as e:
            raise RuntimeError(f"Failed to list drives of {site_name} [Error]:{e}")
    
    def get_site_lists(self, site_name: str, verbose: bool = False) -> pd.DataFrame:
        """
        List all lists for a given SharePoint site.

        Args:
            site_name (str): Name of the SharePoint site.
            verbose (bool, optional): If True, print list details. Defaults to False.

        Returns:
            pd.DataFrame: DataFrame containing list metadata.
        """
        all_list = []
        try:
            site_id = self._get_site_id(f"{self.tenant}/sites/{site_name}/",verbose=verbose)
            lists = requests.get(f"{self.graph}/sites/{site_id}/lists", headers=self.headers).json()
            for l in lists.get('value',None):
                if verbose:
                    print(f"{'-'*20}\n{l}\n{'-'*20}")
                all_list.append({
                    "id":l.get('id',None),
                    "list_name":l.get('name',None),
                    "weburl":l.get('webUrl',None),
                    "created_at":l.get('createdDateTime',None),
                    "created_by": l.get('createdBy', None).get('user', None) if l.get('createdBy', None) is not None else None,
                    "modified_at":l.get('lastModifiedDateTime',None),
                    "site_name":site_name
                    })
        except Exception as e:
            raise RuntimeError(f"Failed to list content for {site_name} [Error]:{e}")
        return pd.DataFrame(all_list)
    
    def get_site_pages(self, site_name: str, verbose: bool = False) -> pd.DataFrame:
        """
        List all pages for a given SharePoint site.

        Args:
            site_name (str): Name of the SharePoint site.
            verbose (bool, optional): If True, print page details. Defaults to False.

        Returns:
            pd.DataFrame: DataFrame containing page metadata.
        """
        try:
            list_item = []
            site_id = self._get_site_id(f"{self.tenant}/sites/{site_name}/",verbose=verbose)
            resp = requests.get(f"{self.graph}/sites/{site_id}/pages", headers=self.headers) #pages,lists,drives
            resp = resp.json()

            # resp.raise_for_status()
            lists = resp["value"]
            for d in resp["value"]:
                if verbose:
                    print(f"List Item : \n{d['name']} | ID: {d['id']} | weburl: {d.get('webUrl',None)}")
                list_item.append({
                                'page_id':d.get('id',None),
                                'page_name': d.get('name',None),
                                "weburl": d.get('webUrl',None),
                                "page_layout": d.get('pageLayout',None),
                                "content_type": d.get("contentType",None).get("name",None),
                                "created_at": d.get('createdDateTime',None),
                                "modified_at": d.get('lastModifiedDateTime',None),                            
                                "created_by": d.get("createdBy",None).get("user",None),
                                "publishing_status":d.get("publishingState",None),
                                'site_name':site_name})
            return pd.DataFrame(list_item)
        except Exception as e:
            raise RuntimeError(f"Failed to list content for {site_name} [Error]:{e}")
    
# ========================================: Get Data Methods :==============================================    
    
    def get_lists_data(self, site_name: str, list_name: str, verbose: bool = False) -> Optional[Dict[str, Any]]:
        """
        Retrieve data fields from a specific list in a SharePoint site.

        Args:
            site_name (str): Name of the SharePoint site.
            list_name (str): Name of the list.
            verbose (bool, optional): If True, print item fields. Defaults to False.

        Returns:
            Optional[Dict[str, Any]]: Dictionary of data fields from the list, or None if not found.
        """
        try:
            site_id = self._get_site_id(f"{self.tenant}/sites/{site_name}/",verbose=verbose)
            lists = requests.get(f"{self.graph}/sites/{site_id}/lists", headers=self.headers).json()
            for li in lists.get('value', []):
                if list_name.lower()==li.get('name').lower():
                    id = li['id']
                    url = f"{self.graph}/sites/{site_id}/lists/{id}/items?$expand=fields"
                    while url:
                        items = requests.get(url, headers=self.headers).json()
                        for item in items.get('value',[]):
                            data_fields = item['fields']
                        if verbose:
                                print(json.dumps(item['fields'],indent=4))
                        url = items.get('@odata.nextLink')
            return data_fields
        except Exception as e:
            raise RuntimeError(f"Failed to generate list content for {site_name} [Error]:{e}")

    def get_site_documents(self, site_name: str, verbose: bool = False, log: bool = False) -> pd.DataFrame:
        """
        Retrieve all documents from all drives in a SharePoint site.

        Args:
            site_name (str): Name of the SharePoint site.
            verbose (bool, optional): If True, print progress and details. Defaults to False.
            log (bool, optional): Unused parameter for logging. Defaults to False.

        Returns:
            pd.DataFrame: DataFrame containing document metadata.
        """
        try:
            doc_df = []  # Initialize doc_df locally at the start of the function
            site_id = self._get_site_id(f"{self.tenant}/sites/{site_name}/",verbose=verbose)
            try:
                self._ensure_token(verbose=verbose)
                drives = requests.get(f"{self.graph}/sites/{site_id}/drives", headers=self.headers).json()
            except Exception as e:
                raise ValueError(f"Error fetching drives in {site_name}: {e}")
                return pd.DataFrame()  # Return empty DataFrame in case of failure
            for d in drives.get('value', []):
                drive_id = d['id']
                if verbose:
                    print(f"Drive ID found: {drive_id} for drive {d['name']}")

                def fetch_all_file(drive_id: str, folder: str = 'root') -> List[Dict[str, Any]]:
                    """
                    Recursively fetch all files in a drive folder.

                    Args:
                        drive_id (str): Drive ID.
                        folder (str, optional): Folder path. Defaults to 'root'.

                    Returns:
                        List[Dict[str, Any]]: List of file metadata dictionaries.
                    """
                    all_files = []
                    url = f"{self.graph}/drives/{drive_id}/{folder}/children?$top=999"
                    while url:
                        try:
                            resp = requests.get(url, headers=self.headers).json()
                        except Exception as e:
                            print(f"Error fetching files: {e}")
                            return all_files  # Return whatever was fetched so far

                        for i, item in enumerate(resp.get('value', [])):
                            parent_path = item['parentReference']['path']
                            folder_Name = parent_path.split('/')[-1]
                            if 'folder' in item:
                                sub_folder_id = f"items/{item['id']}"
                                if verbose:
                                    print(f'Sub-folder found: {sub_folder_id}')
                                all_files.extend(fetch_all_file(drive_id, sub_folder_id))
                            else:
                                all_files.append({
                                    'Doc_Name': item['name'],
                                    'Doc_id': item['id'],
                                    'webUrl': item['webUrl'],
                                    'Created_at':datetime.strptime(item['createdDateTime'], "%Y-%m-%dT%H:%M:%SZ").strftime("%Y-%m-%d %H:%M:%S"),
                                    'Modified_at':pd.to_datetime(datetime.strptime(item['lastModifiedDateTime'], "%Y-%m-%dT%H:%M:%SZ").strftime("%Y-%m-%d %H:%M:%S"), utc=True),
                                    'Doc_Size': f'{round(item['size'] / (1024 * 1024), 2)} MB',
                                    'Folder_Name':folder_Name,
                                    'Drive_Name': d['name']
                                })            
                        self._ensure_token(verbose=verbose)
                        url = resp.get('@odata.nextLink')
                    return all_files

                doc_df.extend(fetch_all_file(drive_id))
                if verbose:
                    print('-' * 10, f"Completed for drive {d['name']}", '-' * 10)
            return pd.DataFrame(doc_df)
        except Exception as e:
            raise RuntimeError(f"Failed to get Documents from {site_name} [Error]:{e}")

    def get_page_canvas(self, site_name: str, page_name: str, verbose: bool = False) -> Dict[str, Any]:
        """
        Retrieve the canvas layout of a specific page in a SharePoint site.

        Args:
            site_name (str): Name of the SharePoint site.
            page_name (str): Name of the page.
            verbose (bool, optional): If True, print progress. Defaults to False.

        Returns:
            Dict[str, Any]: Dictionary containing the canvas layout.
        """
        try:
            site_id = self._get_site_id(f"{self.tenant}/sites/{site_name}/",verbose=verbose)
            page_content = []
            pages = requests.get(f"{self.graph}/sites/{site_id}/pages/", headers=self.headers).json()
            for p in pages['value']:
                page = p['name']
                if p['name'].lower() == page_name.lower():
                    page_id = p['id']
                    if verbose:
                        print(f"Page ID found: {page_id} for page {page}")
                    break
            resp = requests.get(f"{self.graph}/sites/{site_id}/pages/{page_id}/microsoft.graph.sitePage?$expand=canvasLayout", headers=self.headers).json()
            canvas_layout = resp.get('canvasLayout',{})
            return canvas_layout
        except Exception as e:
            raise RuntimeError(f"Failed to get page canvas for {site_name} [Error]:{e}")
    
    def download_file_from_url(self, sp_url: str, verbose: bool = False) -> io.BytesIO:
        """
        Download a SharePoint file given its full URL and return as a BytesIO stream.

        Args:
            sp_url (str): Full SharePoint file URL.
            verbose (bool, optional): If True, print progress. Defaults to False.

        Returns:
            io.BytesIO: In-memory file stream of the downloaded file.
        """
        try:
            self._ensure_token(verbose=verbose)
            parts = urlparse(sp_url)
            query_params = parse_qs(parts.query)
            site_id = self._get_site_id(sp_url,verbose=verbose) 
            if not site_id:
                raise ValueError(f"Site ID not found for URL {sp_url}")

            drives = requests.get(f"{self.graph}/sites/{site_id}/drives", headers=self.headers).json()

            for d in drives['value']:
                if verbose:
                    print(f"Drive Name:{d['name']} || Drive ID: {d['id']}\n",'-'*20)
            
            file_path = parts.path
            if query_params:
                file_keys = ['file','filename','File','FileName']
                for key in file_keys:
                    if key in query_params:
                        file_name=query_params[key][0]
            else:
                file_name = unquote(os.path.basename(file_path))
            extension = file_name.split('.')[-1]
            def find_file_recursive(drive_id: str, folder: str = "root") -> Optional[str]:
                """
                Recursively search for a file in a drive and return its file ID.

                Args:
                    drive_id (str): Drive ID.
                    folder (str, optional): Folder path. Defaults to "root".

                Returns:
                    Optional[str]: File ID if found, else None.
                """
                self._ensure_token(verbose=verbose)
                url = f"{self.graph}/drives/{drive_id}/{folder}/children?$top=999"
                while url:
                    resp = requests.get(url, headers=self.headers).json()
                    for item in resp.get("value", []):
                        if item['name'] == file_name:
                            if verbose:
                                print(f"File Name: {item['name']} | Extension: {extension} | Doc size: {round(item['size']/(1024*1024),2)} MB")
                            return item['id']
                        if "folder" in item:  # go deeper
                            sub_folder = f"items/{item['id']}"
                            found = find_file_recursive(drive_id, sub_folder)
                            if found:
                                return found
                    url = resp.get('@odata.nextLink')

                return None  # Return None when the loop ends without finding the file

            file_id, drive_id = None, None

            # Try each drive until the file is found
            for d in drives['value']:
                if verbose:
                    print(f"Searching in drive: {d['name']}")
                drive_id_try = d['id']
                file_id_try = find_file_recursive(drive_id_try)
                if file_id_try:
                    drive_id, file_id = drive_id_try, file_id_try
                    if verbose:
                        print(f"Found in drive '{d['name']}' with file_id {file_id}")
                    break
            if verbose:
                print(f"File ID found: {file_id}")

            # --- Download the file into temporary memory ---
            resp = requests.get(f"{self.graph}/drives/{drive_id}/items/{file_id}/content", headers=self.headers, stream=True)
            resp.raise_for_status()

            # Use BytesIO to store the content in memory
            file_stream = io.BytesIO()
            for chunk in resp.iter_content(chunk_size=1 << 20):  # 1 MB chunks
                if chunk:
                    file_stream.write(chunk)

            file_stream.seek(0)  
            if verbose:
                print(f"Downloaded file to memory: {file_stream}")
            return file_stream

        except Exception as e:
            raise ValueError('-' * 50, f"[ERROR]: Error while downloading the file: {e}", '-' * 50, '\n\n')

# ==========================================: Get Metadata Methods :====================================
      
    def get_document_metadata(self, site_name: str, list_name: str, verbose: bool = False) -> pd.DataFrame:
        """
        Retrieve metadata for documents in a specific list of a SharePoint site.

        Args:
            site_name (str): Name of the SharePoint site.
            list_name (str): Name of the list.
            verbose (bool, optional): If True, print item details. Defaults to False.

        Returns:
            pd.DataFrame: DataFrame containing document metadata.
        """
        metadata = []
        site_id = self._get_site_id(f"{self.tenant}/sites/{site_name}/",verbose=verbose)
        lists_in = requests.get(f"{self.graph}/sites/{site_id}/lists", headers=self.headers).json()
        for li in lists_in.get('value', []):
            if list_name.lower()==li.get('name').lower():
                id = li['id']
                url = f"{self.graph}/sites/{site_id}/lists/{id}/items?$expand=fields"
                while url:
                    items = requests.get(url, headers=self.headers).json()
                    for index,item in enumerate(items.get('value',[])):
                        if verbose:
                            print(f"item at index {index}:\n{json.dumps(item,indent=4)}\n","-"*100)
                        metadata.append({
                            'ID':item.get('id'),
                            'Review Date':item['fields'].get('Review_x0020_Date'),
                            'Created':item['fields'].get('Created'),
                            'Name':item['fields'].get('LinkFilename'),
                            'weburl':item.get('webUrl'),
                            'Title':item['fields'].get('Title'),
                            'Document Type':item.get('fields').get('Product_x0020_Document_x0020_Type'),
                            'Description':item['fields'].get('KpiDescription'),
                            'Digital':item['fields'].get('Digital'),
                            'Content Language':item['fields'].get('Content_x0020_Language')[0] if isinstance(item['fields'].get('Content_x0020_Language'), list) else None,
                            'Content Type':item.get('fields').get('ContentType'),
                            'PAC/SubPAC Description':item.get('fields').get('PAC_x002F_SubPAC_x0020_Description'),
                            'Modified':item['fields'].get('Modified'),
                            'Opportunity/Problem Area':item['fields'].get('Opportunity_x002F_Problem_x0020_Area'),
                            # 'Modified By':items['fields'].get('Review_x0020_Date'),
                            'Content Owner':item['fields'].get('Content_x0020_Owner_x0028_s_x0029_')[0]['LookupValue'] if isinstance(item['fields'].get('Content_x0020_Owner_x0028_s_x0029_'), list) else None,
                            'Geographical Region':item['fields'].get('Region')[0]if isinstance(item['fields'].get('Region'), list) else None,
                        #     'Link to a Document':item['fields'].get('Review_x0020_Date'),
                            'Primary KC':item['fields'].get('Primary_x0020_Knowledge_x0020_Center',{}).get('Label'),
                            'Primary KC Section':item['fields'].get('Primary_x0020_Knowledge_x0020_Center_x0020_Section',{}).get('Label'),
                            'Primary KC Subsection':item['fields'].get('Primary_x0020_Knowledge_x0020_Center_x0020_Subsection',{}).get('Label'),
                             'Secondary Knowledge Center Section':[item["Label"] for item in item['fields'].get("Secondary_x0020_Knowledge_x0020_Center_x0020_Section_x0028_s_x0029_", []) if "Label" in item],
                            'Publish Status': item['fields'].get('Published_x0020_Status'),
                        #     'Approval Status':item['fields'].get('Review_x0020_Date'),
                            'Application':item['fields'].get('Application'),
                            'Brand':item['fields'].get('Brand')
                            })
                    url = items.get('@odata.nextLink')
        return pd.DataFrame(metadata)
    
    
    def get_page_metadata(self, site_name: str, verbose: bool = False) -> pd.DataFrame:
        """
        Retrieve metadata for all pages in a SharePoint site.

        Args:
            site_name (str): Name of the SharePoint site.
            verbose (bool, optional): If True, print progress. Defaults to False.

        Returns:
            pd.DataFrame: DataFrame containing page metadata.
        """
        page_metadata = []
        site_id = self._get_site_id(f"{self.tenant}/sites/{site_name}/",verbose=verbose)
        url = f"{graph}/sites/{site_id}/lists/Site Pages/items?expand=fields"
        while url:
            page = requests.get(url, headers=self.headers).json()#columns
            for iterator in page['value']:
                page_metadata.append({
                    'ID':iterator.get('id'),
                    'Result Type':iterator.get('fields').get('ContentType'),
                    'Name':iterator.get('fields').get('FileLeafRef'),
                    'Title':iterator.get('fields').get('Title'),
                    'Description':iterator.get('fields').get('Description'),
                    'Content Owner':iterator.get('fields').get('Content_x0020_Owner_x0028_s_x0029_'),
                    'Publish Status':iterator.get('fields').get('Published_x0020_Status'),
                    'Primary Knowledge Center':iterator.get('fields').get('Primary_x0020_Knowledge_x0020_Center'),
                    'Secondary Knowledge Center':iterator.get('fields').get('Secondary_x0020_Knowledge_x0020_Center_x0028_s_x0029_'),
                    'Geographical Region':iterator.get('fields').get('Region'),
                    'Industry/Market':iterator.get('fields').get('Industry_x0020_or_x0020_Market'),
                    'opportunity problem/Area':iterator.get('fields').get('Opportunity_x002F_Problem_x0020_Area'),
                    'Application':iterator.get('fields').get('Application'),
                    'Brand':iterator.get('fields').get('Brand'),
                    'PAC/Sub-PAC Description':iterator.get('fields').get('PAC_x002F_SubPAC_x0020_Description'),
                    'Digital':iterator.get('fields').get('Digital'),
                    'Type':iterator.get('fields').get('DocIcon'),
                    'Created At':iterator.get('fields').get('Created'),
                    'Modified At':iterator.get('fields').get('Modified')
                    })
            url = page.get('@odata.nextLink')
        return pd.DataFrame(page_metadata)

# ===========================================: Protected Methods :==============================================    

    def _get_site_id(self, url: str, verbose: bool = False) -> str:
        """
        Given a SharePoint URL, return the site-id via Graph.

        Args:
            url (str): SharePoint site URL.
            verbose (bool, optional): If True, print progress. Defaults to False.

        Returns:
            str: Site ID.
        """
        parts = urlparse(url)
        host = parts.hostname
        path_parts = parts.path.split("/")
        # Extract site path: e.g. /sites/Finance
        try:
            self._ensure_token(verbose=verbose)
            site_index = path_parts.index("sites")
            site_path = "/sites/" + path_parts[site_index + 1]
            if verbose:
                print('Generated Site_path, started getting site id')
        except ValueError:
            raise ValueError(f"Cannot parse site path from {url}")

        site = requests.get(
            f"{self.graph}/sites/{host}:{site_path}?$select=id,webUrl",
            headers=self.headers, timeout=30
        ).json()
        if site['id']:
            if verbose:
                print("Site Id Found!")
            return site['id']
        else:
            raise ValueError(f"Not able to hit URL check for token value: {url}")

    def _get_headers(self, verbose: bool = False) -> None:
        """
        Generate and set the authentication headers for Microsoft Graph API.

        Args:
            verbose (bool, optional): If True, print progress. Defaults to False.

        Raises:
            AuthenticationError: If authentication with Azure AD fails
        """
        try:
            if verbose:
                print("Generating Token Header")

            app = ConfidentialClientApplication(
                self.client_id,
                authority=self.authority,
                client_credential=self.client_secret
            )

            token_response = app.acquire_token_for_client(scopes=self.scope)

            # Check if we got an error response
            if 'error' in token_response:
                error_msg = token_response.get('error_description', token_response.get('error'))
                raise AuthenticationError(
                    f"Failed to acquire token: {error_msg}",
                    details={
                        'error': token_response.get('error'),
                        'client_id': self.client_id
                    }
                )

            access_token = token_response.get('access_token')
            if not access_token:
                raise AuthenticationError(
                    "No access token in response",
                    details={'response_keys': list(token_response.keys())}
                )

            self.headers = {
                "Authorization": f'Bearer {access_token}',
                'Accept': 'application/json'
            }
            self.token_last_fetch = time.time()

            if verbose:
                print('-' * 10, f'Fetched header at: {self.token_last_fetch}', '-' * 10)

        except AuthenticationError:
            # Already wrapped, just re-raise
            raise
        except Exception as e:
            # Wrap unexpected authentication errors
            raise AuthenticationError(
                f"Unexpected error during authentication: {e}",
                details={
                    'client_id': self.client_id,
                    'authority': self.authority
                }
            ) from e


    def _ensure_token(self, verbose: bool = False) -> None:
        """
        Ensure that the authentication token is valid and refresh if expired.

        Args:
            verbose (bool, optional): If True, print progress. Defaults to False.
        """
        # Removed global keywords; use instance variables directly
        if self.headers is None or time.time() - self.token_last_fetch > (55 * 60):
            if verbose:
                print('-------header Not found, generating new---------')
            self._get_headers(verbose=verbose)
        else:
            remaining_time = 3600 - (time.time() - self.token_last_fetch)
            if verbose:
                print(f'Time left for next Token call: {round(remaining_time / 60, 2)} mins')