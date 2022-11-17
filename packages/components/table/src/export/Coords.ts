export class Coords {
  private _x: number
  private _y: number
  private _value: string

  private constructor(
    lstHeader: Array<Array<string>>,
    row: number,
    col: number
  ) {
    this._x = row
    this._y = col
    this._value = lstHeader[col][row]
  }

  public static of(
    lstHeader: Array<Array<string>>,
    row: number,
    col: number
  ): Coords {
    return new Coords(lstHeader, row, col)
  }

  get x(): number {
    return this._x
  }

  set x(value: number) {
    this._x = value
  }

  get y(): number {
    return this._y
  }

  set y(value: number) {
    this._y = value
  }

  get value(): string {
    return this._value
  }

  set value(value: string) {
    this._value = value
  }
}
