"""
Chart modification agent using LangChain and GPT-4
Handles natural language commands to modify chart styling
"""

from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_core.messages import HumanMessage, AIMessage, SystemMessage
from langchain_core.output_parsers import JsonOutputParser
from typing import Dict, List, Any
import json


class ChartModifierAgent:
    """Agent that interprets natural language commands to modify chart styling"""
    
    def __init__(self, api_key: str, model: str = "gpt-4o"):
        """
        Initialize the chart modifier agent
        
        Args:
            api_key: OpenAI API key
            model: Model name (default: gpt-4o)
        """
        self.llm = ChatOpenAI(
            model=model,
            api_key=api_key,
            temperature=0.1
        )
        
        self.system_prompt = """You are a helpful chart styling assistant. You help users modify Excel chart styling by interpreting their natural language requests.

Your job is to convert user requests into a JSON object with specific styling parameters.

Available parameters to modify:
1. "font_family": Font name (e.g., "Roboto", "Arial", "Calibri", "Times New Roman")
2. "font_size": Font size in points (e.g., 7, 8, 10, 12)
3. "axis_line_color": Hex color for axis lines (e.g., "#BFBFBF", "#000000", "#808080")
4. "axis_color": Hex color for axis text (e.g., "#404040", "#000000")
5. "primary_color": Hex color for primary series (e.g., "#1f77b4", "#FF0000", "#00FF00")
6. "secondary_color": Hex color for secondary series (e.g., "#ff7f0e")
7. "grid": Boolean to show/hide gridlines (true or false)

Common color names to hex mappings:
- gray/grey: "#808080", light gray: "#D3D3D3", dark gray: "#404040"
- black: "#000000", white: "#FFFFFF"
- red: "#FF0000", green: "#00FF00", blue: "#0000FF"
- orange: "#FFA500", yellow: "#FFFF00", purple: "#800080"

IMPORTANT RULES:
1. Only include parameters that the user wants to change
2. Do NOT include parameters that weren't mentioned
3. Always use hex color codes (with #)
4. Font names should be capitalized properly
5. Return ONLY a valid JSON object, nothing else

Examples:

User: "Change the bar color to gray"
Response: {"primary_color": "#808080"}

User: "Make the font Arial size 10"
Response: {"font_family": "Arial", "font_size": 10}

User: "Turn off gridlines and make axis lines black"
Response: {"grid": false, "axis_line_color": "#000000"}

User: "Change primary bars to red and secondary to blue"
Response: {"primary_color": "#FF0000", "secondary_color": "#0000FF"}

User: "Use Calibri font"
Response: {"font_family": "Calibri"}

Now respond to the user's request:"""

        self.parser = JsonOutputParser()
        
        self.prompt = ChatPromptTemplate.from_messages([
            ("system", self.system_prompt),
            MessagesPlaceholder(variable_name="history"),
            ("human", "{input}")
        ])
        
        self.chain = self.prompt | self.llm | self.parser
        
        self.conversation_history: List[Any] = []
    
    def process_command(self, user_input: str, current_style: Dict) -> Dict:
        """
        Process a natural language command and return style modifications
        
        Args:
            user_input: Natural language command from user
            current_style: Current style configuration
            
        Returns:
            Dict with style updates to apply
        """
        try:
            # Invoke the chain
            result = self.chain.invoke({
                "input": user_input,
                "history": self.conversation_history
            })
            
            # Add to conversation history
            self.conversation_history.append(HumanMessage(content=user_input))
            self.conversation_history.append(AIMessage(content=json.dumps(result)))
            
            # Keep only last 10 messages to avoid context overflow
            if len(self.conversation_history) > 10:
                self.conversation_history = self.conversation_history[-10:]
            
            return result
            
        except Exception as e:
            print(f"Error processing command: {e}")
            return {}
    
    def reset_conversation(self):
        """Reset conversation history"""
        self.conversation_history = []
