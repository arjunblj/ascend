(module
  (memory (export "memory") 1)

  (func (export "sum_f64") (param $ptr i32) (param $len i32) (result f64)
    (local $i i32)
    (local $acc v128)
    (local $sum f64)
    (local.set $acc (f64x2.splat (f64.const 0)))
    (block $pairs_done
      (loop $pairs
        (br_if $pairs_done
          (i32.ge_u
            (i32.add (local.get $i) (i32.const 1))
            (local.get $len)
          )
        )
        (local.set $acc
          (f64x2.add
            (local.get $acc)
            (v128.load
              (i32.add
                (local.get $ptr)
                (i32.shl (local.get $i) (i32.const 3))
              )
            )
          )
        )
        (local.set $i (i32.add (local.get $i) (i32.const 2)))
        (br $pairs)
      )
    )
    (local.set $sum
      (f64.add
        (f64x2.extract_lane 0 (local.get $acc))
        (f64x2.extract_lane 1 (local.get $acc))
      )
    )
    (block $tail_done
      (loop $tail
        (br_if $tail_done (i32.ge_u (local.get $i) (local.get $len)))
        (local.set $sum
          (f64.add
            (local.get $sum)
            (f64.load
              (i32.add
                (local.get $ptr)
                (i32.shl (local.get $i) (i32.const 3))
              )
            )
          )
        )
        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $tail)
      )
    )
    (local.get $sum)
  )

  (func (export "min_f64") (param $ptr i32) (param $len i32) (result f64)
    (local $i i32)
    (local $acc v128)
    (local $min f64)
    (if (result f64) (i32.eqz (local.get $len))
      (then (f64.const 0))
      (else
        (if (i32.gt_u (local.get $len) (i32.const 1))
          (then
            (local.set $acc (v128.load (local.get $ptr)))
            (local.set $i (i32.const 2))
            (block $pairs_done
              (loop $pairs
                (br_if $pairs_done
                  (i32.ge_u
                    (i32.add (local.get $i) (i32.const 1))
                    (local.get $len)
                  )
                )
                (local.set $acc
                  (f64x2.min
                    (local.get $acc)
                    (v128.load
                      (i32.add
                        (local.get $ptr)
                        (i32.shl (local.get $i) (i32.const 3))
                      )
                    )
                  )
                )
                (local.set $i (i32.add (local.get $i) (i32.const 2)))
                (br $pairs)
              )
            )
            (local.set $min
              (f64.min
                (f64x2.extract_lane 0 (local.get $acc))
                (f64x2.extract_lane 1 (local.get $acc))
              )
            )
          )
          (else
            (local.set $min (f64.load (local.get $ptr)))
            (local.set $i (i32.const 1))
          )
        )
        (block $tail_done
          (loop $tail
            (br_if $tail_done (i32.ge_u (local.get $i) (local.get $len)))
            (local.set $min
              (f64.min
                (local.get $min)
                (f64.load
                  (i32.add
                    (local.get $ptr)
                    (i32.shl (local.get $i) (i32.const 3))
                  )
                )
              )
            )
            (local.set $i (i32.add (local.get $i) (i32.const 1)))
            (br $tail)
          )
        )
        (local.get $min)
      )
    )
  )

  (func (export "max_f64") (param $ptr i32) (param $len i32) (result f64)
    (local $i i32)
    (local $acc v128)
    (local $max f64)
    (if (result f64) (i32.eqz (local.get $len))
      (then (f64.const 0))
      (else
        (if (i32.gt_u (local.get $len) (i32.const 1))
          (then
            (local.set $acc (v128.load (local.get $ptr)))
            (local.set $i (i32.const 2))
            (block $pairs_done
              (loop $pairs
                (br_if $pairs_done
                  (i32.ge_u
                    (i32.add (local.get $i) (i32.const 1))
                    (local.get $len)
                  )
                )
                (local.set $acc
                  (f64x2.max
                    (local.get $acc)
                    (v128.load
                      (i32.add
                        (local.get $ptr)
                        (i32.shl (local.get $i) (i32.const 3))
                      )
                    )
                  )
                )
                (local.set $i (i32.add (local.get $i) (i32.const 2)))
                (br $pairs)
              )
            )
            (local.set $max
              (f64.max
                (f64x2.extract_lane 0 (local.get $acc))
                (f64x2.extract_lane 1 (local.get $acc))
              )
            )
          )
          (else
            (local.set $max (f64.load (local.get $ptr)))
            (local.set $i (i32.const 1))
          )
        )
        (block $tail_done
          (loop $tail
            (br_if $tail_done (i32.ge_u (local.get $i) (local.get $len)))
            (local.set $max
              (f64.max
                (local.get $max)
                (f64.load
                  (i32.add
                    (local.get $ptr)
                    (i32.shl (local.get $i) (i32.const 3))
                  )
                )
              )
            )
            (local.set $i (i32.add (local.get $i) (i32.const 1)))
            (br $tail)
          )
        )
        (local.get $max)
      )
    )
  )
)
