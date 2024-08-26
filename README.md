# Document Similarity Analysis with Vertex AI

This project utilizes Vertex AI's text embedding capabilities to analyze and compare the similarity between documents in DOCX and RTF formats within a specified directory.

## Installation

1. **Clone the Repository**

   ```bash
   git clone https://your-repository-url.git
   cd document-similarity-analysis

 2. **Create a Virtual Environment**
    
   python -m venv venv
source venv/bin/activate  # On Windows, use `venv\Scripts\activate`

3.**Install dependencies**
   pip install -r requirements.txt

4. **Set up Environment Variables**

Create a .env file in the project root directory.   

Add the following lines, replacing placeholders with your actual GCP project and region:

GCP_PROJECT_ID=your-gcp-project-id
GCP_REGION=your-gcp-region 

## Instructions to run 

**Prepare Your Documents**
Place all your DOCX and RTF files in a single directory.

**Run the Script**
      python document_similarity.py
      Use code with caution.

The script will prompt you to enter the directory path containing your documents.
Alternatively, you can hardcode the directory path within the script (see commented line).

**View Results**

The script will print a similarity matrix to the console, showing pairwise similarity scores between documents.
A CSV file named similarity_matrix_YYYYMMDD_HHMMSS.csv will also be saved in the project directory, containing the same matrix for further analysis.
