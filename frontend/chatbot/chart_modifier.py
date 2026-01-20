"""
Chart modification agent using LangChain + OpenAI
Interprets natural language commands to modify chart styling
"""

from typing import Dict, List, Any
import json
import os

from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_core.messages import HumanMessage, AIMessage
from langchain_core.output_parsers import JsonOutputParser


class ChartModifierAgent:
    """
    Agent that interprets natural language chart styling commands
    and converts them into structured JSON style modifications.
    """

    def __init__(self, model: str = "gpt-4o"):
        # Read API key from environment
        api_key = os.getenv("OPENAI_API_KEY")

        if not api_key:
            raise ValueError(
                "OPENAI_API_KEY not found. "
                "Please set it in a .env file or environment variable."
            )

        self.llm = ChatOpenAI(
            model=model,
            api_key=api_key,
            temperature=0.1
        )

        self.system_prompt = """You are a chart styling assistant.

                                You translate user chart styling instructions into a JSON object.

                                Available parameters:
                                - "font_family"
                                - "font_size"
                                - "axis_line_color"
                                - "axis_color"
                                - "primary_color"
                                - "secondary_color"
                                - "grid"
                                - "chart_title"

                                Rules:
                                1. Only include parameters explicitly requested
                                2. Do NOT invent values
                                3. Always return valid JSON
                                4. Use hex colors with #
                                5. Return NOTHING except the JSON object

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

                                User: "Add title Sales Overview"
                                Response: {{"chart_title": "Sales Overview"}}

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
        try:
            result = self.chain.invoke(
                {
                    "input": user_input,
                    "history": self.conversation_history
                }
            )

            self.conversation_history.append(HumanMessage(content=user_input))
            self.conversation_history.append(AIMessage(content=json.dumps(result)))

            if len(self.conversation_history) > 10:
                self.conversation_history = self.conversation_history[-10:]

            return result if isinstance(result, dict) else {}

        except Exception as e:
            print(f"[ChartModifierAgent] Error: {e}")
            return {}

    def reset_conversation(self):
        self.conversation_history = []