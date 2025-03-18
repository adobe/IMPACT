## IMPACT: Input Method Automation for CJK


IMPACT is an Input Method Editor (IME) automation tool designed to test multilingual typing, rendering, and various other scenarios, specifically for CCJK (Chinese, Simplified Chinese, Japanese, and Korean) locales
Unlike alphabetic languages, these scripts involve thousands of characters (logograms) and require specialized input methods, such as phonetic-based or stroke-based typing.
 
For example:
 
- Chinese uses pinyin or stroke input, where users type based on pronunciation or stroke order.
- Japanese combines three writing systems (Kanji, Hiragana, and Katakana), requiring the IME to switch seamlessly between them.
- Korean uses Hangul, a phonetic alphabet where characters are composed of multiple letters grouped into syllabic blocks.



These complexities introduce challenges in rendering, font handling, and ensuring accurate input conversion. Testing these manually is time-consuming and prone to errors



IMPACT automates this process, including the automatic switching of system keyboards to simulate real-world usage scenarios. allowing comprehensive testing of multilingual typing, rendering, and user interaction scenarios.


By automating these tasks, IMPACT ensures that products are globally compatible, culturally sensitive, and capable of handling the intricacies of CCJK scripts, making it indispensable for developers and testers focused on internationalization.


``` 
CCJK Test Data : testData/testData.json
```

### Execution Guidelines : 
After executiing the Python file. the application will launch automatically. It supports two distinct modes: Live and Workflow, enabling users to interact dynamically or follow a structured process based on their specific requirements.

1. Live Mode:
- We need to traverse to dialog in which we want to test IME and
press “m” key in the dialog.
- After pressing the key “m” Framework will change the IME and
execute the Test data of selected locale in that dialog
- Once all the Test cases executed It will display “press esc to
generate report” message as a indicator in the dialog.
- To generate the report of tested dialogs please press “esc” key.

2. Workflows Mode:
- On clicking Create Workflow It will launch the selected application, we need to make a workflow of dialogs that we want to test to add the dialog in the workflow.
- Please press “r” key on the dialogs to add the dialog in workflow.
- To Save the workflow please press “esc” key. It will redirect to the
UI with the workflow path.
- To execute the workflow, click on Execute Workflow and select
your desired locale. Framework will change the IME and execute
the selected locale CCJK Cases in the dialog.
- Once CCJK Test data executed Framework will generate the
detailed report along with Video and Screenshots
