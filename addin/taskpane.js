Office.onReady(() => {
  document.getElementById("generate").onclick = generateReply;
  document.getElementById("insert").onclick = insertReply;
});

const API_BASE_URL = "https://kontorcompaniet-ai.vercel.app";

async function getEmailData() {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;

    item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
      resolve({
        subject: item.subject || "",
        from: item.from?.emailAddress || "",
        body: bodyResult.status === Office.AsyncResultStatus.Succeeded ? bodyResult.value : "",
      });
    });
  });
}

async function generateReply() {
  const output = document.getElementById("output");
  output.value = "Genererer svar...";

  try {
    const email = await getEmailData();

    const res = await fetch(`${API_BASE_URL}/api/ai/reply`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(email),
    });

    const data = await res.json();
    output.value = data.reply || "Ingen svar mottatt";
  } catch (error) {
    output.value = "Feil ved henting av AI-svar";
  }
}

function insertReply() {
  const text = document.getElementById("output").value;

  Office.context.mailbox.item.body.setSelectedDataAsync(
    text,
    { coercionType: Office.CoercionType.Text }
  );
}