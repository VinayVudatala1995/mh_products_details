# Node.js Excel Processing Tool

## Project Overview

This tool automates the retrieval and update of product data from a RapidAPI endpoint based on product descriptions and packaging specified in an Excel file. It appends product titles, descriptions, and image URLs to the Excel sheet and downloads corresponding images to a local directory.

## Features

- Reads product data from an Excel file.
- Fetches additional product details using RapidAPI.
- Updates the Excel file with fetched data.
- Downloads product images to a local directory.

## Prerequisites

- Node.js (v14.0 or higher recommended)
- npm (Node Package Manager)

## Getting Started

### Installation

Clone the repository to your local machine using:

```bash
git clone https://github.com/yourusername/yourrepositoryname.git
cd yourrepositoryname
```

```
Setup
Install the required npm packages:
npm install
```

Replace key in .env file
RAPIDAPI_KEY=your_rapidapi_key_here

Run command

node index.js
