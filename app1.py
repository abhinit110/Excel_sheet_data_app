import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

def load_data(file_path):
    try:
        df = pd.read_excel(file_path,
                           sheet_name='Sheet1',
                           usecols=['Title', 'Problem', 'Resolved by', 'Comment', 'CL Number'],
                           skiprows=2,
                           engine='openpyxl')
        return df
    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")
        return None

def main():
    st.title("Excel Data Viewer")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        data = load_data(uploaded_file)
        if data is not None:
            st.success("File loaded successfully!")
            st.write("### Data")
            st.dataframe(data)
            st.write("### Column Names")
            st.write(data.columns.tolist())
            st.write("### Summary Statistics")
            st.write(data.describe())

            # Identify categories based on 'Title'
            voc = data['Title'].astype(str).str.contains('VOC', case=True, na=False)
            mr = data['Title'].astype(str).str.contains('MR', case=True, na=False)
            others = ~(voc | mr)

            # Count number of rows in each category
            count_voc = voc.sum()
            count_mr = mr.sum()
            count_others = others.sum()

            st.write(f"### Categories")
            st.write(f"VOC PLM's: {count_voc}")
            st.write(f"MR PLM's: {count_mr}")
            st.write(f"Others: {count_others}")

            # Define a function to calculate the analysed, solved, and CL counts
            def calculate_counts(category):
                solved = data[category]['Resolved by'].astype(str).str.contains('/data protocol/', case=False, na=False)
                analysed = data[category]['Comment'].astype(str).str.contains('/data protocol/', case=False, na=False)
                cl_not_empty = data[category]['CL Number'].notna()
                
                analysed_only = analysed & ~solved
                count_analysed = analysed_only.sum()
                count_solved = solved.sum()
                cl_count = (solved & cl_not_empty).sum()
                
                return count_analysed, count_solved, cl_count

            # Calculate counts for each category
            voc_analysed, voc_solved, voc_cl_count = calculate_counts(voc)
            mr_analysed, mr_solved, mr_cl_count = calculate_counts(mr)
            others_analysed, others_solved, others_cl_count = calculate_counts(others)

            # Display counts for each category
            st.write(f"### Analysis for VOC PLM's")
            st.write(f"Analysed: {voc_analysed}")
            st.write(f"Solved: {voc_solved}")
            st.write(f"Solved with CL Number: {voc_cl_count}")

            st.write(f"### Analysis for MR PLM's")
            st.write(f"Analysed: {mr_analysed}")
            st.write(f"Solved: {mr_solved}")
            st.write(f"Solved with CL Number: {mr_cl_count}")

            st.write(f"### Analysis for Others")
            st.write(f"Analysed: {others_analysed}")
            st.write(f"Solved: {others_solved}")
            st.write(f"Solved with CL Number: {others_cl_count}")

            # Create histogram for each category
            fig, ax = plt.subplots(1, 3, figsize=(12, 4), sharey=True)
            
            categories = ['VOC', 'MR', 'Others']
            analysed_counts = [voc_analysed, mr_analysed, others_analysed]
            solved_counts = [voc_solved, mr_solved, others_solved]
            cl_counts = [voc_cl_count, mr_cl_count, others_cl_count]

            for i, category in enumerate(categories):
                counts = [analysed_counts[i], solved_counts[i], cl_counts[i]]
                bars = ax[i].bar(['Analysed', 'Solved', 'CL counts'], counts, color=['blue', 'green', 'orange'])
                ax[i].set_title(f'{category} PLM\'s', fontsize=10)
                ax[i].tick_params(axis='x', labelsize=6)
                ax[i].tick_params(axis='y', labelsize=6)
                ax[i].yaxis.set_major_locator(plt.MaxNLocator(integer=True))  # Use only natural numbers as indexes
                ax[i].set_ylabel('Count', fontsize=8)
                
                # Add count labels on top of the bars
                for bar in bars:
                    height = bar.get_height()
                    ax[i].text(bar.get_x() + bar.get_width() / 2.0, height, f'{int(height)}', ha='center', va='bottom', fontsize=8)

            st.pyplot(fig)

if __name__ == "__main__":
    main()
