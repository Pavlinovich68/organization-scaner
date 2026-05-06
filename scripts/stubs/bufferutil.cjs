function mask(source, maskBuffer, output, offset, length) {
  for (let index = 0; index < length; index += 1) {
    output[offset + index] = source[index] ^ maskBuffer[index & 3];
  }
}

function unmask(buffer, maskBuffer) {
  for (let index = 0; index < buffer.length; index += 1) {
    buffer[index] ^= maskBuffer[index & 3];
  }
}

module.exports = {
  mask,
  unmask,
};
