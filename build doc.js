const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  HeadingLevel, LevelFormat, PageNumber, NumberFormat
} = require('docx');
const fs = require('fs');

const bulletConfig = {
  reference: "bullets",
  levels: [{
    level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
  }]
};

function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing: { after: 160 },
    children: [new TextRun({ text, font: "Times New Roman", size: 24, bold: opts.bold || false })]
  });
}

function h(text, level = HeadingLevel.HEADING_2) {
  return new Paragraph({
    heading: level,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, font: "Times New Roman", size: 24, bold: true })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 100 },
    children: [new TextRun({ text, font: "Times New Roman", size: 24 })]
  });
}

function center(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: opts.after || 120 },
    children: [new TextRun({ text, font: "Times New Roman", size: opts.size || 24, bold: opts.bold || false })]
  });
}

function blank(space = 200) {
  return new Paragraph({ spacing: { after: space }, children: [new TextRun("")] });
}

const doc = new Document({
  numbering: { config: [bulletConfig] },
  styles: {
    default: { document: { run: { font: "Times New Roman", size: 24 } } },
    paragraphStyles: [
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Times New Roman" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      // Cover Page
      blank(600),
      center("Addis Ababa University", { bold: true, size: 28, after: 200 }),
      center("College of Technology and Built Environment", { bold: true, after: 200 }),
      blank(300),
      center("Group Assignment: Chapter 3 Summary", { bold: true, after: 100 }),
      center("Ethical Decision Making and Moral Judgments", { bold: true, after: 200 }),
      blank(200),
      center("Course Name: MORAL AND CITIZENSHIP EDUCATION (MCED 1011)", { bold: true, after: 100 }),
      center("Section: 12", { bold: true, after: 300 }),
      blank(200),
      center("Group Members", { bold: true, after: 160 }),
      center("Yanet Tilahun ................... UGR/2250/18"),
      center("Yeabsra Workneh ............... UGR/8607/18"),
      center("Yishak Kalu .................... UGR/1868/18"),
      center("Yohanan Temesgen .............. UGR/4349/18"),
      center("Yostena Solomon ............... UGR/5344/18"),
      center("Zebibe Riyad .................. UGR/3911/18"),
      center("Yohannes Debebe ............... UGR/1945/18"),
      center("Yididya Abebe ................. UGR/0143/18"),
      blank(300),
      center("Instructor: Dr. Andergachew Tilahun", { bold: true, after: 120 }),
      center("Submission Date: April 26, 2026", { after: 400 }),

      // Table of Contents
      blank(200),
      h("Table of Contents", HeadingLevel.HEADING_1),
      p("Introduction ........................................................ 3"),
      p("How can we make ethical decisions and actions? ................... 3"),
      p("Ethical Principles and Values of Moral Judgments ................. 4"),
      p("Moral Intuitions and Critical Reasoning .......................... 4"),
      p("Rationalization .................................................. 4"),
      p("Types of Reasoning ............................................... 5"),
      p("Ethics and Religious Faith ....................................... 5"),
      p("Testing Moral Arguments .......................................... 5"),
      p("Thinking Ethically: A Framework for Moral Decision Making ........ 6"),
      p("Fairness and Justice Approach .................................... 6"),
      p("The Common Good Approach ......................................... 6"),
      p("The Rights Approach .............................................. 7"),
      p("To Whom or What Does Morality Apply? ............................. 8"),
      p("Who is Morally/Ethically Responsible? ........................... 9"),
      p("Moral Judgments .................................................. 10"),
      p("What Makes an Action Moral? ..................................... 10"),
      p("Why Morality Matters in Society and Life ........................ 11"),
      p("Conclusion ...................................................... 13"),
      p("References ...................................................... 14"),
      blank(200),

      // Introduction
      h("Chapter Introduction"),
      p("Every day, people are faced with choices about what they want and what they should do. Figuring out what is truly right or wrong is not always straightforward, and that is exactly why morality exists — to offer some guidance when things get complicated. Beyond personal choices, morality also plays a role in managing conflicts between people, giving us shared rules that allow us to cooperate and live together without constant friction."),
      p("Because problems in society never fully go away, ethical theories have developed over time to help us think through moral questions and guide how we act. This chapter looks at ethical decision-making and why, at the end of the day, morality matters."),
      blank(100),
      p("The main goals of this summary are to explore:", { bold: false }),
      bullet("The moral foundations on which we build our ethical standards"),
      bullet("How those standards get applied to specific, real-life situations"),
      bullet("What goes into making a good ethical decision"),
      bullet("Why being moral is something worth caring about"),
      blank(150),

      // How can we make ethical decisions
      h("How Can We Make Ethical Decisions and Actions?"),
      p("Doing the right thing is rarely easy. Our sense of what is \"right\" or \"wrong\" is often shaped by deeply held personal values, but that does not mean all judgments are equal. There are actually two kinds of \"good\" worth distinguishing. The first is called instrumental good — something that is valuable not in itself, but because it helps us get something else we want. Money is a clear example: on its own it is just paper or numbers on a screen, but it becomes useful when we use it to buy food, shelter, or other things that genuinely matter to us. The second type is intrinsic good — things we value simply for what they are, not for what they give us. The same logic applies in reverse: some things are instrumentally bad, and others are intrinsically bad."),
      p("One of the most important tasks in ethical reasoning is to stop and ask yourself: \"Why exactly do I think this is right?\" When it comes to difficult issues like death and dying, for instance, three key values tend to pull in different directions — the sanctity of life, the quality of life, and a person's autonomy, meaning their right to decide for themselves how they want to live and die."),
      p("A second task is checking whether the reasons behind our actions are actually based on solid evidence or sound logic. Ethics does not exist to give people orders or hand them a checklist. Its real purpose is to give us better tools for thinking. Good ethical reasoning is not just about sorting things into \"right\" and \"wrong\" — it is about sitting with the gray areas and taking them seriously."),
      blank(150),

      // Ethical Principles
      h("Ethical Principles and Values of Moral Judgments"),
      p("Ethics is fundamentally about the rules and principles we believe people ought to live by. The problem is that when we try to apply our everyday sense of right and wrong to specific situations, the answers are rarely obvious. To get past that, we need clearer frameworks — structured ways of thinking that help us reach more consistent, logical conclusions."),
      blank(150),

      // Moral Intuitions
      h("Moral Intuitions and Critical Reasoning"),
      p("Part of what ethics does is help us make sense of our gut feelings about what seems \"right\" or \"good.\" Most people have some capacity for empathy — the ability to feel, at least to some degree, what another person is experiencing. This empathy gives rise to what we can call moral sentiments, and when we reason carefully about those sentiments, we develop moral principles. The blend of feeling and reasoning is what forms our conscience. A conscience is rooted in emotion, but it also needs reason to be trustworthy."),
      p("Every society shapes its members through its own ethical ideas, customs, and sometimes formal laws. These forces influence how people feel and what they think is acceptable. Philosophical ethics asks us to step back from all of that and examine our attitudes more critically — not just accepting what we feel or what society tells us, but asking whether those sentiments hold up under scrutiny."),
      blank(150),

      // Rationalization
      h("Rationalization"),
      p("Most of us have some instinct about right and wrong. But philosophy pushes us to go further — to ask not just \"what do I think?\" but \"why do I think it?\" In moral philosophy, an argument is really just a set of reasons offered in support of a belief or action. The goal of ethical debate is not to \"win\" but to think carefully and justify our positions honestly."),
      p("Rationalization is when we use seemingly good reasons to cover up our real motives. It is the difference between genuinely reasoning through a problem and constructing an excuse after the fact. Recognizing this tendency in ourselves is part of what makes ethical thinking hard — and necessary."),
      blank(150),

      // Types of Reasoning
      h("Types of Reasoning"),
      p("Critical reasoning supports arguments through three main forms: analogy, deduction, and induction. These approaches matter because arguments need solid, reliable backing to be convincing."),
      p("Reasoning by analogy works by comparing situations that are similar and drawing conclusions from that comparison. Deductive reasoning starts with a general principle and applies it to a specific case. Inductive reasoning works the other way — it looks at specific evidence and uses it to build toward a general conclusion. Together, these three methods give us a structured way to test and defend our ethical claims, rather than relying purely on what \"feels\" right."),
      blank(150),

      // Ethics and Religious Faith
      h("Ethics and Religious Faith"),
      p("Ethical arguments tend to come from one of two places: religious faith or rational reflection. This creates a real tension, because faith-based morality is grounded in divine command — what God or the scriptures say is right — while philosophical ethics insists on justifying moral claims through reason alone. Religious arguments are still relevant and widely held, but in academic ethics, moral principles generally need to be examined and justified using logic and evidence. This does not mean dismissing faith, but it does mean not letting it go unexamined."),
      blank(150),

      // Testing Moral Arguments
      h("Testing Moral Arguments"),
      p("Critical reasoning involves questioning the assumptions behind any argument, not just accepting them at face value. People develop ethical views through a mix of personal experience, intuition, and reasoning — but those views need to be tested before they can be trusted. There are three main ways to do this:"),
      p("Factual Accuracy: The philosopher David Hume (1711–1776) argued that you cannot derive an \"ought\" from an \"is\" — meaning moral conclusions cannot be pulled directly from facts alone. Still, if the facts an argument depends on are wrong, the whole argument falls apart. Getting the facts right is a necessary first step.", { bold: false }),
      p("Consistency: A moral argument needs to apply the same standards in similar situations. If someone argues that the debts of poorer nations should be cancelled but that personal debts should not, they need to explain what moral difference justifies the distinction. Without that explanation, the argument is inconsistent and loses credibility.", { bold: false }),
      p("Good Will: This is the trickiest one, because good will cannot be easily measured or verified from the outside. It relies more on intuition and character than on logic. Even a technically valid argument carries more weight when the person making it is genuinely motivated by doing what is right.", { bold: false }),
      blank(150),

      // Thinking Ethically
      h("Thinking Ethically: A Framework for Moral Decision Making"),
      p("The starting point for working through any moral issue is getting the facts right. Many ethical disagreements come down to missing or wrong information, so accurate facts matter. That said, facts alone are never enough — they tell us what is, not what ought to be. Resolving moral questions always requires both facts and values working together."),
      p("Because ethical decisions do not always point to a clear answer, thinkers have developed several approaches to help guide judgment. The main ones are fairness and justice, the common good, utilitarianism, rights, and virtue."),
      blank(150),

      // Fairness
      h("Fairness and Justice Approach"),
      p("The fairness or justice approach goes back to Aristotle, who argued that equals should be treated equally. This framework looks at whether an action is truly fair — whether it hands out benefits or burdens in a way that is consistent and free from bias. Favoritism, which gives unjustified benefits to some, and discrimination, which unfairly burdens others, are both morally wrong under this view. The key question this approach asks is: are we distributing benefits and burdens in a way that is consistent and justifiable, or are we treating similar cases differently without good reason?"),
      blank(150),

      // Common Good
      h("The Common Good Approach"),
      p("The common good approach is built on the idea that community life has value in itself. Our individual well-being is not separate from the well-being of the people around us — it is tied to it. This approach calls for compassion and respect, especially toward the most vulnerable members of society, and it emphasizes the shared conditions that make life possible for everyone: things like laws, healthcare systems, public safety, and education."),
      p("The idea has roots going back to Plato, Aristotle, and Cicero, and was developed further by John Rawls. At its core, it asks us to make sure that social systems and institutions serve all members of society fairly — not just those with resources or power. Examples of common good concerns include environmental protection, access to justice, public health, and community safety."),
      p("Notably, this approach overlaps in some ways with the fairness approach — both are concerned with the fair distribution of benefits and burdens among all those affected. The difference is that the common good approach puts more emphasis on shared values and communal relationships, not just equal treatment."),
      blank(150),

      // Rights Approach
      h("The Rights Approach"),
      p("The Rights Approach, which draws heavily on the work of Immanuel Kant, starts from the premise that human beings have inherent dignity. Because people are capable of making free choices, they have a moral right to have those choices respected. This means people must never be treated as mere objects or as means to someone else's ends."),
      p("This approach recognizes several specific rights that deserve protection:"),
      bullet("The Right to the Truth: People have a right to honest information, especially when that information affects important decisions in their lives."),
      bullet("The Right to Privacy: People have the right to live, believe, and speak as they choose in their personal lives, as long as they do not infringe on the rights of others."),
      bullet("The Right Not to Be Injured: People have the right not to be harmed unless they have freely and knowingly done something that justifies punishment, or willingly accepted a risk."),
      bullet("The Right to What Is Agreed: If someone has made a promise or entered into a contract freely, others have a right to expect that promise to be honored."),
      p("When applying this framework, the central question is: does this action respect the moral rights of everyone involved? The more seriously an action violates someone's rights, the more morally wrong it is."),
      blank(150),

      // To Whom Does Morality Apply
      h("To Whom or What Does Morality Apply?"),
      p("Morality applies to individuals, groups, societies, and organizations — basically, to any entity capable of making choices and being held responsible for them. It is the set of values and rules that guide behavior and help people judge right from wrong. There are four main dimensions to consider:"),
      p("Religious Morality: This dimension concerns the relationship between humans and a higher power. In many traditions, right and wrong are defined by religious teachings, commandments, or holy texts. Violating these rules may be seen as a moral failure toward God, even if it does not directly harm other people. For example, in Islamic theology, the concept of Tawhid — the absolute belief in one God — is central. Violating it through shirk (associating others with God) is considered the gravest sin. On a more everyday level, someone might donate to charity because their faith commands generosity, or avoid dishonesty because they believe God sees all actions.", { bold: false }),
      p("Morality and Nature: This dimension looks at how humans relate to the natural world — plants, animals, ecosystems. The idea here is that nature is not just a resource for humans to use; it has its own value. Acting morally toward nature means caring for it for its own sake, not just because a healthy environment benefits us. This kind of moral thinking underlies things like tree planting to prevent soil erosion, recycling, keeping public spaces clean, and the creation of parks and wildlife sanctuaries.", { bold: false }),
      p("Individual Morality: Sometimes a person acts morally — or refuses to act immorally — not because of religious rules or social expectations, but simply because their own conscience tells them it is wrong. This internal sense of right and wrong is individual morality. It matters because it shows that ethical behavior does not always need external enforcement. A person might choose to return lost money even when no one would ever know. They might speak up against an injustice even when it costs them something. These choices come from within.", { bold: false }),
      p("Social Morality: This is arguably the most important dimension because it directly affects other people and society as a whole. Social morality is about how we treat each other — it is grounded in shared norms, cultural values, and the principles that keep a community functioning: respect, fairness, responsibility, and care for others. Examples include respecting others' opinions, not harming others, treating elders with dignity, and upholding the kinds of norms that hold communities together. Without social morality, societies would fragment quickly.", { bold: false }),
      blank(150),

      // Who is Responsible
      h("Who is Morally/Ethically Responsible?"),
      p("Moral responsibility, as we generally understand it, applies to human beings. Holding supernatural beings morally accountable requires faith. Holding animals or plants accountable is complicated by the fact that their behavior is largely instinctual — not the result of deliberate choice."),
      p("That said, recent research into animal cognition is interesting. Some experiments have shown that certain animals can develop thought processes that look surprisingly human-like. It is possible that, in the future, our understanding of moral agency could expand to include some animals. For now, though, the scientific evidence suggests that most animals, and certainly plants, are non-moral beings — they act without the capacity to weigh right from wrong."),
      p("The clearest illustration of this is the way we respond to harm. If a wolf kills our chickens, we might take action to protect the flock — but we do not put the wolf on trial. If a person commits the same act of theft, we take them to court. The reason is simple: humans are the ones capable of understanding what is right and wrong and making choices accordingly. That is the foundation of moral responsibility."),
      blank(150),

      // Moral Judgments
      h("Moral Judgments"),
      p("A moral judgment is essentially a decision about whether an action is right or wrong. Importantly, moral judgments apply to voluntary actions — things people choose freely, not things that happen to them or that they do by accident. These judgments are guided by normative standards, meaning shared ideas about how people should behave."),
      p("Moral judgment is rarely simple. Consider a doctor caring for a terminally ill patient who is in severe, unrelenting pain. The patient wants to die. The doctor has taken an oath to preserve life and is bound by law. But the doctor also knows that if they were in the same situation, they might want the same thing. This kind of scenario puts multiple moral principles in direct conflict — and there is no easy answer."),
      p("When we make moral judgments, we typically consider four elements:"),
      bullet("Motives — Why is the action being taken?"),
      bullet("Means — How is it being carried out?"),
      bullet("Consequences — What effects does it have?"),
      bullet("Moral situation — Are there competing values that make a straightforward choice impossible?"),
      blank(150),

      // What Makes Action Moral
      h("What Makes an Action Moral?"),
      p("Not every action has moral weight. Wearing a watch, watering a plant — these are morally neutral. But stealing, lying, or harming someone are clearly morally significant. The question is: what makes the difference?"),
      p("Three features tend to make an action morally relevant:"),
      bullet("Agents: There needs to be an entity with free will — someone capable of choosing."),
      bullet("Intention: The motive behind an action matters. Unintentional harm tends to be morally neutral. But harm caused by negligence — where someone should have known better — is a different matter and does carry moral weight."),
      bullet("Impact on others: Moral actions are those that affect other people in meaningful ways, whether positively or negatively — physically, mentally, or emotionally."),
      blank(150),

      // Why Morality Matters
      h("Why Morality Matters in Society and Life"),
      p("1. Can Society Survive Without Morality?", { bold: true }),
      p("Morality is really important for keeping society peaceful and organized. It is because of values like honesty, fairness, and respect that people can live together and treat each other properly. If we did not have morality, many people would only think about themselves. That could lead to selfishness, crime, and disorder. Laws can help control how people behave, but they are not enough if people do not care about doing what is right. Morality helps people work together, solve problems, and build a community. Without it, society would have a hard time functioning. For example, if everyone lied or stole, people would lose trust and feel unsafe."),
      blank(80),
      p("2. Why Does Trust Depend on Morality?", { bold: true }),
      p("Trust depends on people being honest, responsible, and fair. People trust others when they believe they will tell the truth and do the right thing. If people are dishonest or selfish, trust becomes weak. Without trust, it is hard to have relationships, work together, and build strong communities. For instance, in school, if students cheat or refuse to do their part in group work, the whole group suffers."),
      blank(80),
      p("3. Do All Humans Deserve Respect?", { bold: true }),
      p("Yes, all human beings deserve respect because every person has value and dignity. Respect is a universal right, not a privilege for a few. It involves treating everyone with fairness and dignity, whether you agree with them or not. Without respect, relationships can fall apart and society can become divided. Morality and respect help create peace, understanding, and cooperation among people. For example, respecting someone's culture, opinions, or identity even if they are different from yours is very important."),
      blank(80),
      p("4. Does Being Good Pay Off in the Long Run?", { bold: true }),
      p("Being good usually brings positive results over time. People who are honest, kind, and helpful often build strong relationships and earn a good reputation. Others are more likely to trust and support them. While being good may not always give immediate rewards, it often leads to lasting benefits in the future. For instance, an honest person may face challenges now, but later they gain respect and trust."),
      blank(80),
      p("5. Why Does Doing Good Feel Good?", { bold: true }),
      p("Helping others is a rewarding act that creates joy and purpose. When we do good things, it can make both the person helping and the person receiving help feel better. It also strengthens connections between people. Knowing that you helped someone can bring confidence and inner peace. For example, helping a person in need can make you feel proud and satisfied afterward."),
      blank(150),

      // Conclusion
      h("Conclusion"),
      p("Ethical reasoning is not a luxury — it is something we genuinely need, both as individuals and as members of society. Having a shared vocabulary of ethical principles and frameworks gives us a way to approach difficult moral questions with more clarity and less guesswork."),
      p("For anyone facing an ethical decision, it helps to have some kind of process to follow — a way to make sure you have thought through the relevant factors and considered not just your own interests but those of everyone affected. Without that kind of structure, even well-intentioned decisions can cause harm."),
      p("Aristotle's insight here is worth taking seriously: moral character is not built through knowing the right theory. It is built through practice. You do not become an honest person by reading about honesty. You become one by choosing honesty repeatedly, even when it is inconvenient. And the reason to bother? Because morality is not just an abstract set of rules — it is what makes a good life, and a good society, possible."),
      blank(150),

      // References
      h("References"),
      p("A Framework for Thinking Ethically. ETHICS 1, no. 2 (Winter 1988)."),
      p("MORAL AND CITIZENSHIP EDUCATION (MCED 1011) Course Material."),
      p("Boss, J. A. (1999). Analyzing Moral Issues. Mayfield Publishing Company."),
      p("Thiroux, J. (1995). Ethics: Theory and Practice (5th ed.). Prentice Hall."),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/claude/HUMANIZED_CIVIC_ASSIGNMENT.docx", buffer);
  console.log("Done");
});