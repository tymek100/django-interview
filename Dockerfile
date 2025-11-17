# Use Ubuntu as the base image
FROM ubuntu:22.04

# Avoid interactive prompts
ARG DEBIAN_FRONTEND=noninteractive

# Install system dependencies
RUN apt-get update && apt-get install -y \
    python3 python3-pip python3-venv curl git \
    && apt-get clean

# Install PDM
RUN curl -sSL https://pdm-project.org/install-pdm.py | python3 -

# Add PDM to PATH
ENV PATH="/root/.local/bin:${PATH}"

# Set the working directory
WORKDIR /app

# Copy project files
COPY . /app

# Install dependencies via PDM
RUN pdm install

# Run Django migrations
RUN pdm run python manage.py migrate

# Expose Django port
EXPOSE 8000

# Run Django server
CMD ["pdm", "run", "python", "manage.py", "runserver", "0.0.0.0:8000"]
