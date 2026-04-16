// Destination folder ID where the generated documents will be moved to
const destinationFolderId = 'put_id_of_the_folder_where_you_want_to_receive_generated_Document';

// Configuration for each template document
const templateConfig = {
  initialDocument: {
    fileId: 'put_id_Of_the_template_file_here',
    namePattern: '{Client LAST name} Client Agreement',
    placeholders: {
      'PA Name (LAST, First)': 'PA Name (LAST, First)',
      'A number': 'A number',
      'Client FIRST name': 'Client FIRST name',
      'Client LAST name': 'Client LAST name'
    }
  }
  // Add more template configurations here if needed
};
