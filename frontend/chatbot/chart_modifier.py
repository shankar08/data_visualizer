"""
Chart modification agent using LangChain + OpenAI
Interprets natural language commands to modify chart styling
"""

from typing import Dict, List, Any
import json

from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_core.messages import HumanMessage, AIMessage
from langchain_core.output_parsers import JsonOutputParser


class ChartModifierAgent:
    """
    Agent that interprets natural language chart styling commands
    and converts them into structured JSON style modifications.
    """

    def __init__(self, api_key: str, model: str = "gpt-4o"):
        self.llm = ChatOpenAI(
            model=model,
            api_key=api_key,
            temperature=0.1
        )

        # IMPORTANT:
        # All JSON examples use double braces {{ }} so LangChain
        # does NOT treat them as template variables
        self.system_prompt = """You are a chart styling assistant.

You translate user chart styling instructions into a JSON object.

Available parameters:
- "font_family": string
- "font_size": integer (6–72)
- "axis_line_color": hex color
- "axis_color": hex color
- "primary_color": hex color
- "secondary_color": hex color
- "grid": boolean

Rules:
1. Only include parameters explicitly requested
2. Do NOT invent values
3. Always return valid JSON
4. Use hex colors (with #)
5. Return NOTHING except the JSON object

Color examples:
- gray: "#808080"
- light gray: "#D3D3D3"
- dark gray: "#404040"
- black: "#000000"
- red: "#FF0000"
- blue: "#0000FF"

Examples:

User: Change bar color to gray
Response: {{ "primary_color": "#808080" }}

User: Use Arial font size 10
Response: {{ "font_family": "Arial", "font_size": 10 }}

User: Remove gridlines
Response: {{ "grid": false }}

User: Make axis labels black
Response: {{ "axis_color": "#000000" }}

User: Change primary to red and secondary to blue
Response: {{ "primary_color": "#FF0000", "secondary_color": "#0000FF" }}

Now respond to the user's command.
"""

        self.prompt = ChatPromptTemplate.from_messages(
            [
                ("system", self.system_prompt),
                MessagesPlaceholder(variable_name="history"),
                ("human", "{input}")
            ]
        )

        self.parser = JsonOutputParser()

        self.chain = self.prompt | self.llm | self.parser

        self.conversation_history: List[Any] = []

    def process_command(self, user_input: str, current_style: Dict) -> Dict:
        """
        Convert natural language input into style modifications

        Args:
            user_input: User instruction (e.g. "Add title Hello")
            current_style: Current chart style (unused, but future-proof)

        Returns:
            Dict of style updates
        """
        try:
            result = self.chain.invoke(
                {
                    "input": user_input,
                    "history": self.conversation_history
                }
            )

            # Update conversation memory
            self.conversation_history.append(
                HumanMessage(content=user_input)
            )
            self.conversation_history.append(
                AIMessage(content=json.dumps(result))
            )

            # Keep memory small & fast
            if len(self.conversation_history) > 10:
                self.conversation_history = self.conversation_history[-10:]

            # Guarantee dict output
            if isinstance(result, dict):
                return result

            return {}

        except Exception as e:
            print(f"[ChartModifierAgent] Error: {e}")
            return {}

    def reset_conversation(self):
        """Clear conversation history"""
        self.conversation_history = []