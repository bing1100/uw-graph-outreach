# Email template

Edit the subject line and the text between the two `---` lines. Anything in `{{double braces}}` is filled per person from the CSV column of the same name, or from `--var` for values that are the same for everyone (your name, program, school). A row with an empty placeholder is skipped and listed, so no email goes out half-filled.

**Subject:** ATTN: {{FullName}}, {{Title}} @ {{Organization}} – {{Topic}}

---

Hi {{FirstName}},

My name is {{SenderName}} and I'm studying {{SenderProgram}} at the University of Waterloo. I'm currently investigating {{Topic}}.

My understanding is that {{Problem}}. I came across {{Personalization}}, which is why I thought of you.

As part of my investigation, I'm trying to validate a potential solution: {{Solution}}. Your expertise would be incredibly helpful in proving or disproving my hypothesis!

Would you be open to connecting for ~15 minutes on {{Slot}}?

Best,
{{SenderName}}

---

## Tips

- **Subject.** The "ATTN: Name, Title @ Organization – topic" line tells the reader at a glance that the email is meant for them. Keep the topic to a few words, such as "Bottlenecks in acute care diagnostics".
- **Problem.** Write it as your current understanding ("My understanding is that…"), not a claim. Experts respond to being asked to confirm or correct something specific.
- **Personalization.** One real detail per person, such as a paper, a talk or their role, does more than any amount of general praise. Write it so it follows "I came across", for example "your paper on emergency department flow".
- **Slot.** A specific time ("Wednesday or Thursday at 2 pm ET") gets more replies than "whenever works". Use the recipient's time zone when they're far away.
- **Opt-out line.** Canada's anti-spam law (CASL) expects a way to opt out of follow-ups. Keep the last line, and honour opt-outs straight away.
- **Plain text.** The script sends plain text, so formatting such as bold or links in brackets appears literally. Paste URLs as plain URLs.
