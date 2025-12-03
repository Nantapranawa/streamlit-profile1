# Use an official Python runtime as a base image
FROM python:3.13-slim

# Set the working directory inside the container
WORKDIR /telkom_app

# Copy the local files into the container
COPY . /telkom_app

# Install the dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Expose the port Streamlit will run on
EXPOSE 8501

# Set the command to run your app using Streamlit
CMD ["streamlit", "run", "new.py"]