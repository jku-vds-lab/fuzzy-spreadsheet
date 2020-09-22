import { hsl, rgb } from "d3";
import { floor } from "mathjs";
/* global console */

export default class Bins {
  private minDomain: number;
  private maxDomain: number;
  private width: number;
  private nrOfBins: number;
  public bins: Bin[]

  constructor(minDomain: number, maxDomain: number, width: number) {
    this.minDomain = minDomain;
    this.maxDomain = maxDomain;
    this.width = width;
    this.nrOfBins = floor((this.maxDomain - this.minDomain) / this.width);
    this.bins = new Array<Bin>();
  }

  createBins(data: number[]) {

    let i = 0;
    this.bins = new Array<Bin>();

    while (i < this.nrOfBins) {

      let bin: Bin = { x0: 0, x1: 0, length: 0, samples: new Array<number>() };

      bin.x0 = this.minDomain + this.width * i;
      bin.x1 = bin.x0 + this.width;
      bin.length = 0;
      bin.samples = new Array<number>();
      this.bins.push(bin);
      i++;
    }

    data.forEach((element: number) => {
      this.bins.forEach((bin: Bin) => {
        if (element >= bin.x0 && element < bin.x1) {
          bin.samples.push(element);
          bin.length++;
        }
      });
    })
    return this.bins;
  }

  getTickValues() {
    let ticks = new Array<number>();
    let i = this.minDomain;

    while (i <= this.maxDomain) {
      ticks.push(i);
      i = i + 20;
    }

    return ticks;
  }

  generateBlueColors() {

    let blueColors = [];
    let i = 0;
    let color = rgb(217, 217, 217)
    blueColors.push(color.hex());
    const stepSizeR = 199 / this.nrOfBins;
    const stepSizeG = 233 / this.nrOfBins;
    const stepSizeB = 192 / this.nrOfBins;
    while (i < this.nrOfBins) { // from range: RGB(199,233,192) to RGB(0,68,27)
      color = rgb(199 - stepSizeR * i, 233 - stepSizeG * i, 192 - stepSizeB * i); // in steps of 14
      blueColors.push(color.hex());

      i++;
    }

    console.log('Final color: ' + color);
    return blueColors;
  }

  generateOrangeColors() {

    let orangeColors = [];
    let i = 0;
    let color = rgb(217, 217, 217);
    orangeColors.push(color.hex());

    const stepSizeG = 225 / this.nrOfBins;
    const stepSizeB = 239 / this.nrOfBins;
    while (i < this.nrOfBins) {
      let color = rgb(231, 225 - stepSizeG * i, 239 - stepSizeB * i);
      orangeColors.push(color.hex());
      i++;
    }

    return orangeColors;
  }

  generateRedBlueColors() {

    let redBlueColors = [];
    let redColors = [];
    let blueColors = [];
    let i = 0;

    while (i < 50) {
      let color = rgb(178 + 0.69 * 2 * i, 24 + 2.23 * 2 * i, 43 + 2.04 * 2 * i);
      redColors.push(color.hex());
      i++;
    }

    i = 0;
    while (i <= 50) {
      let color = rgb(247 - 2.14 * 2 * i, 247 - 1.45 * 2 * i, 247 - 0.75 * 2 * i);
      blueColors.push(color.hex());
      i++;
    }

    redBlueColors.push(...redColors);
    redBlueColors.push(...blueColors);
    return redBlueColors;
  }


  static getRedColorsForImpact() {
    let redColors = [];
    let i = 0;

    while (i <= 100) {
      let color = rgb(247 - 0.69 * i, 247 - 2.23 * i, 247 - 2.04 * i);
      redColors.push(color.hex());
      i++;
    }
    return redColors;
  }

  static getBlueColorsForImpact() {
    let blueColors = [];
    let i = 0;

    while (i <= 100) {
      let color = rgb(247 - 2.14 * 2 * i, 247 - 1.45 * 2 * i, 247 - 0.75 * 2 * i);
      blueColors.push(color.hex());
      i++;
    }
    return blueColors;
  }


}

interface Bin {
  x0: number;
  x1: number;
  length: number;
  samples: number[];
}


