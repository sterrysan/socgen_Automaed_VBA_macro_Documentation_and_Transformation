pip install transformers torch
pip install graphviz gpt4all
pip install fpdf2 pillow
import gradio as gr
import re
from gpt4all import GPT4All
import graphviz
from fpdf import FPDF
from PIL import Image

# Initialize the GPT4All model
model_name = "Meta-Llama-3-8B-Instruct.Q4_0.gguf"
model = GPT4All(model_name)

def explain_vba_code(vba_code):
    # Prepare the input prompt
    prompt = f"Explain the following VBA code:\n\n{vba_code}\n\nExplanation:"

    # Generate the explanation within a chat session
    with model.chat_session():
        explanation = model.generate(prompt, max_tokens=250)

    # Extract key elements from the code
    elements = extract_elements(vba_code)

    # Create the process flow diagram
    flow_diagram = create_flow_diagram(elements)
    flow_diagram.render('process_flow_diagram', format='png')

    # Analyze code quality and security
    quality_issues = analyze_code_quality(vba_code)
    security_issues = analyze_security(vba_code)

    quality_text = "Code Quality and Efficiency Analysis:\n"
    quality_text += "\n".join(quality_issues) if quality_issues else "No issues found."

    security_text = "Security Vulnerability Analysis:\n"
    security_text += "\n".join(security_issues) if security_issues else "No vulnerabilities found."

    # Create the PDF document
    pdf = PDF()

    # Add explanation text to the PDF
    pdf.add_page()
    pdf.chapter_title('Explanation')
    pdf.chapter_body(explanation)

    # Add process flow diagram to the PDF
    pdf.add_flow_diagram('process_flow_diagram.png')

    # Add code quality and security analysis to the PDF
    pdf.add_page()
    pdf.chapter_title('Code Quality and Efficiency Analysis')
    pdf.chapter_body(quality_text)

    pdf.add_page()
    pdf.chapter_title('Security Vulnerability Analysis')
    pdf.chapter_body(security_text)

    # Output the PDF to a file
    pdf_file_path = 'VBA_Macro_Documentation.pdf'
    pdf.output(pdf_file_path)

    return explanation, 'process_flow_diagram.png', pdf_file_path

# Define functions used above
def extract_elements(vba_code):
    elements = []
    lines = vba_code.strip().split('\n')
    for line in lines:
        line = line.strip()
        if line.startswith("Sub "):
            func_name = re.match(r'Sub (\w+)', line).group(1)
            elements.append((f"Start of function: {func_name}", "oval"))
        elif line.startswith("Dim "):
            var_name = re.match(r'Dim (\w+)', line).group(1)
            elements.append((f"Declare variable: {var_name}", "rectangle"))
        elif "=" in line and not line.startswith("If "):
            parts = line.split('=')
            var_name = parts[0].strip()
            value = parts[1].strip()
            elements.append((f"{var_name} assigned value: {value}", "rectangle"))
        elif line.startswith("If "):
            condition = re.match(r'If (.+) Then', line).group(1)
            elements.append((f"Decision: If {condition}", "diamond"))
        elif line.startswith("Else"):
            elements.append(("Else", "rectangle"))
        elif line.startswith("End If"):
            elements.append(("End If", "rectangle"))
        elif line.startswith("MsgBox "):
            msg = re.match(r'MsgBox (.+)', line).group(1)
            elements.append((f"Display message: {msg}", "rectangle"))
        elif line.startswith("End Sub"):
            elements.append(("End of function", "oval"))
    return elements

def create_flow_diagram(elements):
    dot = graphviz.Digraph(comment='Process Flow Diagram')
    dot.node('Start', 'Start', shape='oval')

    previous_node = 'Start'
    decision_count = 0

    for i, (element, shape) in enumerate(elements):
        node_name = f"Step{i}"

        if "Decision" in element:
            decision_node = node_name
            dot.node(decision_node, element, shape=shape)
            decision_count += 1
            yes_node = f"DecisionYes{decision_count}"
            no_node = f"DecisionNo{decision_count}"
            dot.node(yes_node, "Yes", shape="oval")
            dot.node(no_node, "No", shape="oval")
            dot.edge(previous_node, decision_node)
            dot.edge(decision_node, yes_node, label="True")
            dot.edge(decision_node, no_node, label="False")
            previous_node = yes_node  # Continue from the "Yes" branch
        elif element == "Else":
            previous_node = no_node  # Continue from the "No" branch
        else:
            dot.node(node_name, element, shape=shape)
            dot.edge(previous_node, node_name)
            previous_node = node_name

    dot.node('End', 'End', shape='oval')
    dot.edge(previous_node, 'End')
    
    return dot

def analyze_code_quality(vba_code):
    issues = []
    lines = vba_code.strip().split('\n')
    
    for i, line in enumerate(lines):
        line = line.strip()
        if len(line) > 80:
            issues.append(f"Line {i+1}: Line length exceeds 80 characters.")
        if 'Goto' in line:
            issues.append(f"Line {i+1}: Avoid using 'Goto' statements.")
        if line.count("'") > 1:
            issues.append(f"Line {i+1}: Multiple comments in a single line.")
    
    return issues

def analyze_security(vba_code):
    vulnerabilities = []
    lines = vba_code.strip().split('\n')
    
    for i, line in enumerate(lines):
        line = line.strip()
        if 'Shell' in line:
            vulnerabilities.append(f"Line {i+1}: Potential shell execution detected.")
        if 'Eval' in line:
            vulnerabilities.append(f"Line {i+1}: Potential use of 'Eval' detected.")
    
    return vulnerabilities

# Create the PDF class
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'VBA Macro Documentation', 0, 1, 'C')

    def chapter_title(self, title):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, 0, 1, 'L')
        self.ln(10)

    def chapter_body(self, body):
        self.set_font('Arial', '', 12)
        self.multi_cell(0, 10, body)
        self.ln()

    def add_flow_diagram(self, image_path):
        self.add_page()
        self.image(image_path, x=10, y=30, w=165, h=250)

# Create the Gradio interface
iface = gr.Interface(
    fn=explain_vba_code,
    inputs=gr.components.Textbox(lines=20, placeholder="Enter your VBA code here..."),
    outputs=[
        gr.components.Textbox(label="Explanation"),
        gr.components.Image(type="filepath", label="Process Flow Diagram"),
        gr.components.File(label="Download PDF")
    ],
    title="VBA Macro Documentation Tool",
    description="Upload your VBA code to get an explanation, a process flow diagram, and a documentation PDF."
)

# Launch the interface
iface.launch()
