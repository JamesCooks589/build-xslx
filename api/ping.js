module.exports = (req, res) => {
  res.status(200).json({ ok: true, now: Date.now(), url: req.url });
};
