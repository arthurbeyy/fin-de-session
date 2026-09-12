const scene = document.querySelector('.machine-wrap');
const machine = document.querySelector('#machine');

if (scene && machine && matchMedia('(pointer: fine)').matches) {
  scene.addEventListener('pointermove', (event) => {
    const box = scene.getBoundingClientRect();
    const x = (event.clientX - box.left) / box.width - 0.5;
    const y = (event.clientY - box.top) / box.height - 0.5;
    machine.style.transform = `rotateX(${3 - y * 8}deg) rotateY(${-3 + x * 10}deg)`;
  });
  scene.addEventListener('pointerleave', () => {
    machine.style.transform = 'rotateX(3deg) rotateY(-3deg)';
  });
}
