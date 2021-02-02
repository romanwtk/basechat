function getMessage() {
    let user = this.context.pageContext.user.displayName;
    let message = this.domElement.querySelector("textarea").value;
    console.log(message);
  }

function innerEmoji(emoji) {
  var txaMessage = document.getElementById('message');
  txaMessage.value += emoji;
}