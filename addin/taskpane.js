Office.onReady(() => {
  document.getElementById("generate").onclick = generateReply;
  document.getElementById("insert").onclick = insertReply;
});

async function getEmailData() {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;

    item.body.getAsync("text", (bodyResult) => {
      resolve({
        subject: item.subject,
        from: item.from?.emailAddress || "",
        body: bodyResult.value,
      });
    });
  });
}

async function generateReply() {
  const output = document.getElementById("output");
  output.value = "Genererer...";

  const email = await getEmailData();

  const res = await fetch("https://localhost:3000/api/ai/reply", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(email),
  });

  const data = await res.json();

  output.value = data.reply || "Ingen svar";
}

function insertReply() {
  const text = document.getElementById("output").value;

  Office.context.mailbox.item.body.setSelectedDataAsync(
    text,
    { coercionType: Office.CoercionType.Text }
  );
}