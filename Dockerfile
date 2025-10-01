    # Use an official Python runtime as a parent image (for Windows containers)
    FROM python:3.9-slim

    # Set the working directory in the container
    WORKDIR /app

    # Copy the requirements.txt file into the container
    COPY requirements.txt .

    # Install any needed packages specified in requirements.txt
    RUN pip install --no-cache-dir -r requirements.txt

    # Copy the current directory contents into the container at /app
    COPY . .

    # # Define environment variable (optional)
    # ENV NAME World

    # Run app.py when the container launches
    CMD ["streamlit","run", "streamlit_new.py", "--server.address=0.0.0.0", "--server.port=8501"]

    EXPOSE 8501