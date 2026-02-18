import express from "express";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

app.post("/agent", async (req, res) => {
    const { prompt } = req.body;

    const agentResponse = await fetch("https://api.foundry.ai/agents/asst_gGwnLmkX0JbtKhY327Ipagi6/run", {
        method: "POST",
        headers: {
            "Authorization": `Bearer ${process.env.FOUNDRY_API_KEY}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify({ input: prompt })
    }).then(r => r.json());

    res.json({ text: agentResponse.output });
});

app.listen(3000, () => console.log("Backend running"));
