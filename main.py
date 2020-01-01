from functions import read_journal_responses, update_journal


def main():
    # Acquire form data
    df = read_journal_responses()

    # Post to Google doc
    update_journal(df)

    return


if __name__ == "__main__":

    main()