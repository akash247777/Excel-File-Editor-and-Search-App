import streamlit as st
import pandas as pd
from io import BytesIO

# Streamlit App
def main():
    st.title("Excel File Editor App")

    # Specify the file paths
    file_path = 'C:\\Search\\supplier.xlsx'
    matching_file_path = 'C:\\Search\\itemsearch.xlsx'
    
    @st.cache_data
    def load_data(file_path):
        return pd.read_excel(file_path)
    
    @st.cache_data
    def load_matching_data(matching_file_path):
        return pd.read_excel(matching_file_path)
    
    # Load the Excel files into DataFrames
    if "df" not in st.session_state:
        try:
            st.session_state.df = load_data(file_path)
            st.session_state.matching_df = load_matching_data(matching_file_path)
            st.write("Loaded Excel Files:")
        except Exception as e:
            st.error(f"Error loading data: {e}")
            return

    df = st.session_state.df
    matching_df = st.session_state.matching_df

    # Ensure all columns are strings to avoid AttributeError
    df = df.applymap(lambda x: str(x) if pd.notnull(x) else "")

    # Step 2: Add a new column
    new_column_name = st.text_input("Enter the name of the new column:", "NewColumn")
    
    if new_column_name:
        # Initialize the new column with empty values
        if new_column_name not in df.columns:
            df[new_column_name] = ""  # Add the new column

        # Create two columns for layout
        col1, col2 = st.columns(2)

        with col1:
            # Step 3: Display the DataFrame as an editable table
            st.write("Edit the table below:")
            edited_df = st.data_editor(
                df,
                num_rows="dynamic",
                height=800,  # Set a fixed height for the table
                width=800,   # Set a fixed width for the table
                use_container_width=True,  # Expand to container width
                key="data_editor",  # Unique key for the data editor
            )

            # Save the updated data to the Excel file
            if st.button("Save Changes"):
                try:
                    edited_df.to_excel(file_path, index=False, engine='openpyxl')
                    st.success("Changes saved successfully!")
                except Exception as e:
                    st.error(f"Error saving changes: {e}")

        with col2:
            # Step 4: Filter matching list based on user input in the new column
            for i, row in edited_df.iterrows():
                user_input = row[new_column_name]
                if user_input:
                    # Convert both the user input and itemname to lowercase for case-insensitive comparison
                    matching_items = matching_df[
                        matching_df['itemname'].str.lower().str.startswith(user_input.lower())
                    ]
                    if not matching_items.empty:
                        st.write(f"Matching items for row {i + 1}:")
                        st.dataframe(
                            matching_items,
                            height=800,  # Set a fixed height for the table
                            width=800,   # Set a fixed width for the table
                            use_container_width=True,  # Expand to container width
                        )
                    

        # Step 5: Download the updated Excel file
        st.write("Download the updated file:")
        output = BytesIO()
        edited_df.to_excel(output, index=False, engine='xlsxwriter')
        output.seek(0)
        st.download_button(
            label="Download Excel file",
            data=output,
            file_name="updated_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()