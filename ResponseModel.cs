﻿namespace LLmDocxTrans;

public class ResponseModel
{
    public Choice[]? choices { get; set; }

    public class Choice
    {
        public Message? message { get; set; }
    }

    public class Message
    {
        public string? content { get; set; }
    }
}