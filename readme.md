# SentiXL

SentiXL is a sentiment analysis tool for processing Excel files with subjective responses using the GPT model. It analyzes the sentiment of each response and generates a new Excel file with sentiment labels.

## Key Features

- Automated sentiment analysis of subjective responses in Excel files
- Utilizes GPT model for accurate sentiment classification
- Generates a new Excel file with sentiment labels
- User-friendly interface for easy integration

## Getting Started

1. Install dependencies: `pip install -r requirements.txt`
2. Set up OpenAI API key and system prompt in `.env` file using `.env-sample`
3. Place Excel file in `input` folder
4. Run SentiXL: `python main.py`
5. Get sentiment-labeled Excel file from `output` folder

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

## License

This project is licensed under the [MIT License](LICENSE).